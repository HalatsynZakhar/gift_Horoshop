from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass, replace
from itertools import chain
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


DEFAULT_EXPORT_LIMIT = 500
PUBLIC_LOG_PATH_JSON_PATTERN = re.compile(r'("public_log_path"\s*:\s*")(?P<value>[^"]*)(")')
MAPPING_HEADER_MARKERS = (
    "артикул",
    "article",
    "варіант",
    "вариант",
    "variant",
    "подартикул",
    "підартикул",
    "субартикул",
    "source",
    "target",
    "ключ",
)


class HoroshopGiftsError(RuntimeError):
    pass


@dataclass(frozen=True)
class Settings:
    domain: str
    host: str
    port: int
    batch_size: int
    request_timeout_seconds: int
    public_log_path: Path
    public_log_name: str
    login: str = ""
    password: str = ""
    auth_enabled: bool = True
    auth_login: str = "admin"
    auth_password: str = "Bublik44"
    images_xml_url: str = "http://176.104.6.149:8080/xml/images_export.xml"

    @property
    def public_log_file(self) -> Path:
        return self.public_log_path / self.public_log_name


@dataclass(frozen=True)
class Credentials:
    login: str = ""
    password: str = ""
    token: str = ""


@dataclass(frozen=True)
class GiftRow:
    primary_display: str
    gift_display: str
    row_number: int
    action: str = "upsert"


@dataclass(frozen=True)
class CatalogProduct:
    article: str
    article_for_display: str
    gifts: tuple[Any, ...]


@dataclass(frozen=True)
class GiftPlan:
    primary_display: str
    gift_display: str
    primary_article: str
    gift_article: str
    row_number: int
    action: str = "upsert"
    error: str = ""
    user_primary: str = ""
    user_gift: str = ""
    existing_gifts: tuple[str, ...] = ()
    updated_gifts: tuple[str, ...] = ()

    @property
    def ready(self) -> bool:
        return not self.error


@dataclass(frozen=True)
class GiftAssociation:
    primary_display: str
    gift_display: str
    primary_article: str
    gift_article: str


def normalize(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def leading_zeros_count(text: str) -> int:
    count = 0
    for character in text:
        if character == "0":
            count += 1
        else:
            break
    return count


def normalize_article_code(value: Any) -> str:
    text = normalize(value)
    if not text:
        return ""
    raw_alpha = "".join(character if character.isalnum() else "-" for character in text).strip("-")
    zeros = leading_zeros_count(raw_alpha)
    cleaned = re.sub(r"-{2,}", "-", raw_alpha)
    if zeros > 0 and not cleaned.startswith("0" * zeros):
        cleaned = ("0" * zeros) + cleaned.lstrip("0")
    return cleaned


def header_key(value: Any) -> str:
    return " ".join(normalize(value).casefold().split())


@dataclass(frozen=True)
class ArticleMappings:
    entries: dict[str, tuple[str, ...]]

    @classmethod
    def empty(cls) -> "ArticleMappings":
        return cls({})

    def merged_with(self, other: "ArticleMappings") -> "ArticleMappings":
        combined = {source: list(targets) for source, targets in self.entries.items()}
        for source, targets in other.entries.items():
            combined.setdefault(source, []).extend(targets)
        result: dict[str, tuple[str, ...]] = {}
        for source, targets in combined.items():
            deduped = tuple(dict.fromkeys(targets))
            if any(target.casefold() != source.casefold() for target in deduped):
                result[source] = deduped
        return ArticleMappings(result)

    def targets_for(self, article: str, use_normalized: bool) -> tuple[str, ...]:
        if not self.entries:
            return (article,)
        exact = self.entries.get(article)
        if exact:
            return exact
        folded = article.casefold()
        folded_matches = [targets for source, targets in self.entries.items() if source.casefold() == folded]
        if len(folded_matches) == 1:
            return folded_matches[0]
        if use_normalized:
            norm_article = normalize_article_code(article)
            zeros = leading_zeros_count(norm_article)
            normalized = norm_article.casefold()
            normalized_matches = [
                targets
                for source, targets in self.entries.items()
                if normalize_article_code(source).casefold() == normalized
                and leading_zeros_count(normalize_article_code(source)) == zeros
            ]
            if len(normalized_matches) == 1:
                return normalized_matches[0]
        return (article,)


def mapping_column_indexes(headers: tuple[Any, ...]) -> tuple[int, ...]:
    indexes = tuple(index for index, value in enumerate(headers) if normalize(value))
    if len(indexes) < 2:
        raise ValueError("У файлі сопоставлень має бути мінімум два стовпці з варіантами артикула.")
    return indexes


def is_mapping_header(row: tuple[Any, ...]) -> bool:
    return any(marker in header_key(value) for value in row for marker in MAPPING_HEADER_MARKERS)


def parse_article_mappings(data: bytes) -> ArticleMappings:
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        worksheet = workbook.worksheets[0]
        iterator = worksheet.iter_rows(values_only=True)
        try:
            first_row = next(iterator)
        except StopIteration as error:
            raise ValueError("Файл сопоставлень порожній.") from error
        has_header = is_mapping_header(first_row)
        if has_header:
            column_indexes = mapping_column_indexes(first_row)
            rows = iterator
            start_row = 2
        else:
            column_indexes = tuple(range(len(first_row)))
            rows = chain((first_row,), iterator)
            start_row = 1
        mappings: dict[str, list[str]] = {}
        for row_number, row in enumerate(rows, start=start_row):
            if not row or all(value is None for value in row):
                continue
            if has_header:
                raw_articles = [
                    normalize(row[index] if index < len(row) else None)
                    for index in column_indexes
                    if normalize(row[index] if index < len(row) else None)
                ]
            else:
                raw_articles = [normalize(value) for value in row if normalize(value)]
            unique_articles: list[str] = list(dict.fromkeys(raw_articles))
            if len(unique_articles) < 2:
                continue
            if len({a.casefold() for a in unique_articles}) < 2:
                continue
            for i in range(len(unique_articles)):
                for j in range(i + 1, len(unique_articles)):
                    a, b = unique_articles[i], unique_articles[j]
                    if a.lstrip("0") == b.lstrip("0") and a != b:
                        raise ValueError(
                            f"Рядок {row_number}: артикули «{a}» та «{b}» відрізняються лише початковими нулями (це різні бренди/товари). "
                            f"Їх заборонено зіставляти один з одним."
                        )
            for source in unique_articles:
                mappings.setdefault(source, []).extend(unique_articles)

        final_entries: dict[str, tuple[str, ...]] = {}
        for source, targets in mappings.items():
            deduped = tuple(dict.fromkeys(targets))
            if any(target.casefold() != source.casefold() for target in deduped):
                final_entries[source] = deduped

        if not final_entries:
            raise ValueError(
                "У файлі сопоставлень немає жодного корисного правила (усі рядки порожні, містять лише один артикул або однакові значення)."
            )
        return ArticleMappings(final_entries)
    finally:
        workbook.close()


def load_article_mappings(path: Path) -> ArticleMappings:
    if not path.exists():
        return ArticleMappings.empty()
    try:
        raw_entries = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise HoroshopGiftsError(f"Не вдалося прочитати базу сопоставлень: {path}.") from error
    if not isinstance(raw_entries, dict):
        raise HoroshopGiftsError("База сопоставлень має некоректний формат.")
    entries: dict[str, tuple[str, ...]] = {}
    for source, targets in raw_entries.items():
        source_article = normalize(source)
        if not source_article or not isinstance(targets, list):
            raise HoroshopGiftsError("База сопоставлень має некоректний формат.")
        target_articles = tuple(dict.fromkeys(normalize(target) for target in targets if normalize(target)))
        if not target_articles:
            raise HoroshopGiftsError("База сопоставлень має некоректний формат.")
        if any(target.casefold() != source_article.casefold() for target in target_articles):
            entries[source_article] = target_articles
    return ArticleMappings(entries)


def save_article_mappings(path: Path, mappings: ArticleMappings) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    serialized = {source: list(targets) for source, targets in mappings.entries.items()}
    temporary_path = path.with_suffix(path.suffix + ".tmp")
    try:
        temporary_path.write_text(json.dumps(serialized, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary_path.replace(path)
    except OSError as error:
        raise HoroshopGiftsError(f"Не вдалося зберегти базу сопоставлень: {path}.") from error


def extract_clusters(mappings: ArticleMappings) -> list[list[str]]:
    visited: set[str] = set()
    clusters: list[list[str]] = []
    for source in sorted(mappings.entries.keys()):
        if source in visited:
            continue
        cluster: list[str] = []
        queue = [source]
        visited.add(source)
        while queue:
            current = queue.pop(0)
            cluster.append(current)
            for target in mappings.entries.get(current, ()):
                if target not in visited:
                    visited.add(target)
                    queue.append(target)
        if len(cluster) >= 2:
            clusters.append(sorted(cluster))
    clusters.sort(key=lambda item: item[0])
    return clusters


def endpoint_url(domain: str, endpoint: str) -> str:
    return urljoin(f"{domain.rstrip('/')}/", endpoint.lstrip("/"))


def repair_public_log_path_json(config_text: str) -> str:
    def replace_path(match: re.Match[str]) -> str:
        value = re.sub(r"(?<!\\)\\(?!\\)", lambda _: "\\\\", match.group("value"))
        return f'{match.group(1)}{value}{match.group(3)}'

    return PUBLIC_LOG_PATH_JSON_PATTERN.sub(replace_path, config_text)


def load_settings(config_file: Path) -> Settings:
    config_text = config_file.read_text(encoding="utf-8-sig")
    try:
        raw = json.loads(config_text)
    except json.JSONDecodeError as original_error:
        repaired = repair_public_log_path_json(config_text)
        if repaired == config_text:
            raise original_error
        try:
            raw = json.loads(repaired)
        except json.JSONDecodeError:
            raise original_error
        try:
            config_file.write_text(repaired, encoding="utf-8")
        except OSError:
            pass
    if not isinstance(raw, dict):
        raise ValueError("config.json must contain an object.")

    horoshop = raw.get("horoshop") or {}
    server = raw.get("server") or {}
    logging_config = raw.get("logging") or {}
    auth_config = raw.get("auth") or {}
    if not isinstance(horoshop, dict) or not isinstance(server, dict) or not isinstance(logging_config, dict):
        raise ValueError("Sections horoshop, server and logging must contain objects.")
    domain = normalize(horoshop.get("domain"))
    if not domain:
        raise ValueError("Set horoshop.domain in config.json.")

    public_log_path = Path(normalize(logging_config.get("public_log_path", "logs")) or "logs")
    if not public_log_path.is_absolute():
        public_log_path = config_file.parent / public_log_path
    public_log_name = normalize(logging_config.get("public_log_name", "horoshop_gifts.log")) or "horoshop_gifts.log"
    if Path(public_log_name).name != public_log_name:
        raise ValueError("logging.public_log_name must be a file name without a path.")

    login = normalize(horoshop.get("login"))
    password = normalize(horoshop.get("password"))

    auth_enabled = True
    auth_login = "admin"
    auth_password = "Bublik44"
    if isinstance(auth_config, dict):
        if "enabled" in auth_config:
            auth_enabled = bool(auth_config.get("enabled"))
        if normalize(auth_config.get("login")):
            auth_login = normalize(auth_config.get("login"))
        if normalize(auth_config.get("password")):
            auth_password = normalize(auth_config.get("password"))
    if (not auth_password or auth_password == "Bublik44") and normalize(raw.get("admin_password")):
        auth_password = normalize(raw.get("admin_password"))

    images_xml_url = normalize(raw.get("images_xml_url")) or "http://176.104.6.149:8080/xml/images_export.xml"

    return Settings(
        domain=domain.rstrip("/"),
        host=normalize(server.get("host", "0.0.0.0")) or "0.0.0.0",
        port=max(1, min(65535, int(server.get("port", 8094)))),
        batch_size=max(1, int(horoshop.get("batch_size", 50))),
        request_timeout_seconds=max(1, int(horoshop.get("request_timeout_seconds", 60))),
        public_log_path=public_log_path,
        public_log_name=public_log_name,
        login=login,
        password=password,
        auth_enabled=auth_enabled,
        auth_login=auth_login,
        auth_password=auth_password,
        images_xml_url=images_xml_url,
    )


def parse_action(value: Any) -> str:
    marker = normalize(value).casefold()
    if marker in {"", "ні", "нет", "no", "0"}:
        return "upsert"
    if marker in {"так", "да", "yes", "1", "видалити", "удалить", "delete"}:
        return "delete"
    raise ValueError("у колонці «Видалити (Так)» вкажіть Так або залиште її порожньою.")


def parse_excel_gifts(data: bytes) -> list[GiftRow]:
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        worksheet = workbook.worksheets[0]
        rows: list[GiftRow] = []
        seen: set[tuple[str, str]] = set()
        headers = {"основний товар", "артикул основного товару", "primary article", "основной товар"}
        for row_number, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
            if not row or all(value is None for value in row[:3]):
                continue
            primary = normalize(row[0] if len(row) > 0 else "")
            if primary.casefold() in headers:
                continue
            gift = normalize(row[1] if len(row) > 1 else "")
            try:
                action = parse_action(row[2] if len(row) > 2 else "")
                if not primary:
                    raise ValueError("вкажіть артикул основного товару.")
                if not gift:
                    raise ValueError("вкажіть артикул подарунка.")
                key = (primary.casefold(), gift.casefold())
                if key in seen:
                    raise ValueError("цей зв'язок товару й подарунка повторюється.")
                seen.add(key)
            except ValueError as error:
                raise ValueError(f"Рядок {row_number}: {error}") from error
            rows.append(GiftRow(primary, gift, row_number, action))
        if not rows:
            raise ValueError("Excel не містить жодного зв'язку подарунка.")
        return rows
    finally:
        workbook.close()


def _style_header(worksheet: Any) -> None:
    fill = PatternFill("solid", fgColor="166534")
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")


def build_excel_template() -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Подарунки"
    worksheet.append(["Артикул основного товару", "Артикул подарунка"])
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = "A1:B1"
    worksheet.column_dimensions["A"].width = 36
    worksheet.column_dimensions["B"].width = 36
    _style_header(worksheet)
    for row in range(2, 102):
        worksheet.cell(row=row, column=1).number_format = "@"
        worksheet.cell(row=row, column=2).number_format = "@"

    guide = workbook.create_sheet("Інструкція")
    guide.column_dimensions["A"].width = 105
    guide["A1"] = "Шаблон для призначення подарунків"
    guide["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    guide["A1"].fill = PatternFill("solid", fgColor="166534")
    guide.merge_cells("A1:B1")
    instructions = [
        "Заповнюйте лист «Подарунки». Кожен рядок додає подарунок до основного товару (або всіх його варіацій з бази сопоставлень).",
        "В обох колонках можна вказувати article_for_display (артикул зі сайту), внутрішній article Хорошоп або артикул із бази сопоставлень.",
        "Сервіс знаходить внутрішні артикули в експорті каталогу та передає їх Хорошопу.",
        "Наявні подарунки основного товару не стираються: додається лише новий зв'язок.",
        "Для масового видалення вивантажте поточний список: у ньому є колонка «Видалити (Так)».",
    ]
    for row, instruction in enumerate(instructions, start=3):
        guide.cell(row=row, column=1, value=instruction).alignment = Alignment(wrap_text=True, vertical="top")
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def build_registry_excel(associations: list[GiftAssociation]) -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Подарунки"
    worksheet.append(["Артикул основного товару", "Артикул подарунка", "Видалити (Так)"])
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = "A1:C1"
    worksheet.column_dimensions["A"].width = 36
    worksheet.column_dimensions["B"].width = 36
    worksheet.column_dimensions["C"].width = 18
    _style_header(worksheet)
    for association in associations:
        worksheet.append([association.primary_display, association.gift_display, ""])
    for row in range(2, worksheet.max_row + 1):
        worksheet.cell(row=row, column=1).number_format = "@"
        worksheet.cell(row=row, column=2).number_format = "@"
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def build_mappings_excel(mappings: ArticleMappings) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Таблиця сопоставлень"

    clusters = extract_clusters(mappings)
    max_cols = max((len(item) for item in clusters), default=2)
    max_cols = max(max_cols, 2)

    headers = [f"Артикул {index}" for index in range(1, max_cols + 1)]
    sheet.append(headers)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(row=1, column=len(headers)).coordinate}"
    _style_header(sheet)

    for row_index, cluster in enumerate(clusters, start=2):
        row_values = list(cluster) + [""] * (max_cols - len(cluster))
        sheet.append(row_values)
        for column_index in range(1, max_cols + 1):
            sheet.cell(row=row_index, column=column_index).number_format = "@"

    for column_index in range(1, max_cols + 1):
        column_letter = sheet.cell(row=1, column=column_index).column_letter
        sheet.column_dimensions[column_letter].width = 30

    register_sheet = workbook.create_sheet("Повний реєстр")
    register_headers = ["Пошуковий артикул", "Кількість відповідностей", "Всі відповідні артикули"]
    register_sheet.append(register_headers)
    register_sheet.freeze_panes = "A2"
    register_sheet.auto_filter.ref = "A1:C1"
    _style_header(register_sheet)

    for row_index, (source, targets) in enumerate(sorted(mappings.entries.items()), start=2):
        register_sheet.append([source, len(targets), ", ".join(targets)])
        register_sheet.cell(row=row_index, column=1).number_format = "@"
        register_sheet.cell(row=row_index, column=3).number_format = "@"

    register_sheet.column_dimensions["A"].width = 28
    register_sheet.column_dimensions["B"].width = 24
    register_sheet.column_dimensions["C"].width = 60

    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def build_failed_excel(failures: list[tuple[GiftPlan, str, str]]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Не встановлено"
    headers = [
        "Статус",
        "Помилка",
        "Рядок",
        "Операція",
        "Артикул основного товару (вхідний)",
        "Артикул подарунка (вхідний)",
        "Артикул основного товару (сайт)",
        "Внутрішній артикул основного товару",
        "Артикул подарунка (сайт)",
        "Внутрішній артикул подарунка",
        "Видалити (Так)",
    ]
    sheet.append(headers)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(row=1, column=len(headers)).coordinate}"
    for row_index, (plan, status, message) in enumerate(failures, start=2):
        action_label = "Видалення" if plan.action == "delete" else "Додавання"
        delete_flag = "Так" if plan.action == "delete" else ""
        sheet.append([
            status,
            message,
            plan.row_number,
            action_label,
            plan.user_primary or plan.primary_display,
            plan.user_gift or plan.gift_display,
            plan.primary_display if plan.primary_article else "",
            plan.primary_article,
            plan.gift_display if plan.gift_article else "",
            plan.gift_article,
            delete_flag,
        ])
        for col_idx in (5, 6, 7, 8, 9, 10):
            sheet.cell(row=row_index, column=col_idx).number_format = "@"

    widths = {
        "A": 16,
        "B": 52,
        "C": 10,
        "D": 16,
        "E": 32,
        "F": 32,
        "G": 32,
        "H": 32,
        "I": 32,
        "J": 32,
        "K": 18,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    _style_header(sheet)
    for row in sheet.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def build_report_excel(entries: list[tuple[GiftPlan, str, str]]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Детальний звіт подарунків"
    headers = [
        "Рядок",
        "Операція",
        "Артикул основного товару (вхідний)",
        "Артикул основного товару (сайт)",
        "Внутрішній артикул основного товару",
        "Артикул подарунка (вхідний)",
        "Артикул подарунка (сайт)",
        "Внутрішній артикул подарунка",
        "Подарунки до операції",
        "Подарунки після операції",
        "Статус",
        "Примітки",
    ]
    sheet.append(headers)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(row=1, column=len(headers)).coordinate}"

    for row_index, (plan, status, message) in enumerate(entries, start=2):
        action_label = "Видалення" if plan.action == "delete" else "Додавання"
        sheet.append([
            plan.row_number,
            action_label,
            plan.user_primary or plan.primary_display,
            plan.primary_display,
            plan.primary_article,
            plan.user_gift or plan.gift_display,
            plan.gift_display,
            plan.gift_article,
            ", ".join(plan.existing_gifts),
            ", ".join(plan.updated_gifts),
            status,
            message,
        ])
        for col_idx in (3, 4, 5, 6, 7, 8, 9, 10):
            sheet.cell(row=row_index, column=col_idx).number_format = "@"

    widths = {
        "A": 10,
        "B": 16,
        "C": 32,
        "D": 32,
        "E": 32,
        "F": 32,
        "G": 32,
        "H": 32,
        "I": 38,
        "J": 38,
        "K": 20,
        "L": 46,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    _style_header(sheet)
    for row_cells in sheet.iter_rows(min_row=2):
        row_cells[-1].alignment = Alignment(wrap_text=True, vertical="top")
        row_cells[-3].alignment = Alignment(wrap_text=True, vertical="top")
        row_cells[-4].alignment = Alignment(wrap_text=True, vertical="top")
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def gift_article(value: Any) -> str:
    if isinstance(value, dict):
        return normalize(value.get("article"))
    return normalize(value)


class CatalogIndex:
    def __init__(self, products: list[CatalogProduct]) -> None:
        self.products = products
        self.by_article = {product.article: product for product in products}
        self.by_display: dict[str, list[CatalogProduct]] = {}
        self.by_display_folded: dict[str, list[CatalogProduct]] = {}
        self.by_exact: dict[str, list[CatalogProduct]] = {}
        self.by_folded: dict[str, list[CatalogProduct]] = {}
        self.by_normalized: dict[str, list[CatalogProduct]] = {}
        for product in products:
            if product.article_for_display:
                self.by_display.setdefault(product.article_for_display, []).append(product)
                self.by_display_folded.setdefault(product.article_for_display.casefold(), []).append(product)
            for key in {product.article, product.article_for_display} - {""}:
                self.by_exact.setdefault(key, []).append(product)
                self.by_folded.setdefault(key.casefold(), []).append(product)
                normalized = normalize_article_code(key)
                if normalized:
                    self.by_normalized.setdefault(normalized.casefold(), []).append(product)

    @classmethod
    def from_raw(cls, raw_products: list[dict[str, Any]]) -> "CatalogIndex":
        products: list[CatalogProduct] = []
        for raw in raw_products:
            article = normalize(raw.get("article")) if isinstance(raw, dict) else ""
            if not article:
                continue
            gifts = raw.get("gifts", [])
            products.append(
                CatalogProduct(
                    article=article,
                    article_for_display=normalize(raw.get("article_for_display")),
                    gifts=tuple(gifts) if isinstance(gifts, list) else (),
                )
            )
        return cls(products)

    def resolve(self, value: str, use_normalized: bool = False) -> tuple[CatalogProduct | None, str]:
        display_exact = list({item.article: item for item in self.by_display.get(value, [])}.values())
        if len(display_exact) == 1:
            return display_exact[0], ""
        if len(display_exact) > 1:
            return None, f"Артикул відображення '{value}' не є унікальним."

        exact = list({item.article: item for item in self.by_exact.get(value, [])}.values())
        if len(exact) == 1:
            return exact[0], ""
        if len(exact) > 1:
            return None, f"Артикул відображення '{value}' не є унікальним."

        display_folded = list({item.article: item for item in self.by_display_folded.get(value.casefold(), [])}.values())
        if len(display_folded) == 1:
            return display_folded[0], ""
        if len(display_folded) > 1:
            return None, f"Артикул відображення '{value}' не є унікальним."

        folded = list({item.article: item for item in self.by_folded.get(value.casefold(), [])}.values())
        if len(folded) == 1:
            return folded[0], ""
        if len(folded) > 1:
            return None, f"Артикул відображення '{value}' не є унікальним."

        if use_normalized:
            norm_val = normalize_article_code(value)
            zeros = leading_zeros_count(norm_val)
            normalized = norm_val.casefold()
            candidates = self.by_normalized.get(normalized, [])
            valid_candidates: list[CatalogProduct] = []
            for item in candidates:
                for key in {item.article, item.article_for_display} - {""}:
                    norm_k = normalize_article_code(key)
                    if norm_k.casefold() == normalized and leading_zeros_count(norm_k) == zeros:
                        valid_candidates.append(item)
                        break
            normalized_matches = list({item.article: item for item in valid_candidates}.values())
            if len(normalized_matches) == 1:
                return normalized_matches[0], ""
            if len(normalized_matches) > 1:
                return None, f"Нормалізований артикул '{norm_val}' для '{value}' не є унікальним."

        return None, f"Артикул відображення '{value}' не знайдений у каталозі."

    def resolve_display(self, value: str, use_normalized: bool = False) -> tuple[CatalogProduct | None, str]:
        return self.resolve(value, use_normalized=use_normalized)

    def display_for_article(self, article: str) -> str:
        product = self.by_article.get(article)
        if product is None:
            return article
        return product.article_for_display or product.article

    def associations(self) -> list[GiftAssociation]:
        pairs: list[GiftAssociation] = []
        for primary in self.products:
            for raw_gift in primary.gifts:
                article = gift_article(raw_gift)
                if not article:
                    continue
                gift = self.by_article.get(article)
                pairs.append(
                    GiftAssociation(
                        primary_display=primary.article_for_display or primary.article,
                        gift_display=(gift.article_for_display or gift.article) if gift else article,
                        primary_article=primary.article,
                        gift_article=article,
                    )
                )
        return sorted(pairs, key=lambda pair: (pair.primary_display.casefold(), pair.gift_display.casefold()))


def _resolve_targets(
    raw_value: str,
    catalog: CatalogIndex,
    normalize_articles: bool,
    mappings: ArticleMappings,
) -> tuple[list[CatalogProduct], str, bool]:
    targets = mappings.targets_for(raw_value, normalize_articles)
    mapped = len(targets) > 1 or (len(targets) == 1 and targets[0] != raw_value)
    products: dict[str, CatalogProduct] = {}
    errors: list[str] = []
    for target in targets:
        product, error = catalog.resolve(target, use_normalized=normalize_articles)
        if product is not None:
            products[product.article] = product
        elif error and "не знайдений у каталозі" not in error:
            errors.append(error)
    if not products:
        if errors:
            return [], errors[0], mapped
        if mapped:
            return [], f"Жоден варіант артикула для '{raw_value}' не знайдений у каталозі.", mapped
        return [], f"Артикул відображення '{raw_value}' не знайдений у каталозі.", mapped
    return list(products.values()), "", mapped


def _format_gift_list(raw_gifts: tuple[Any, ...] | list[Any], catalog: CatalogIndex) -> tuple[str, ...]:
    result: list[str] = []
    for item in raw_gifts:
        art = gift_article(item)
        if art:
            result.append(catalog.display_for_article(art))
    return tuple(result)


def prepare_plan(
    rows: list[GiftRow],
    catalog: CatalogIndex,
    normalize_articles: bool = False,
    mappings: ArticleMappings | None = None,
) -> list[GiftPlan]:
    active_mappings = mappings or ArticleMappings.empty()
    plan: list[GiftPlan] = []
    seen_pairs: set[tuple[str, str, str]] = set()

    for row in rows:
        primaries, primary_error, primary_mapped = _resolve_targets(
            row.primary_display, catalog, normalize_articles, active_mappings
        )
        if primary_error:
            plan.append(
                GiftPlan(
                    primary_display=row.primary_display,
                    gift_display=row.gift_display,
                    primary_article="",
                    gift_article="",
                    row_number=row.row_number,
                    action=row.action,
                    error=primary_error,
                    user_primary=row.primary_display,
                    user_gift=row.gift_display,
                )
            )
            continue

        gifts, gift_error, gift_mapped = _resolve_targets(
            row.gift_display, catalog, normalize_articles, active_mappings
        )
        if gift_error:
            single_primary_article = primaries[0].article if len(primaries) == 1 else ""
            plan.append(
                GiftPlan(
                    primary_display=row.primary_display,
                    gift_display=row.gift_display,
                    primary_article=single_primary_article,
                    gift_article="",
                    row_number=row.row_number,
                    action=row.action,
                    error=gift_error,
                    user_primary=row.primary_display,
                    user_gift=row.gift_display,
                )
            )
            continue

        multi_expanded = len(primaries) > 1 or len(gifts) > 1 or primary_mapped or gift_mapped
        row_plans: list[GiftPlan] = []
        same_product_skipped = False

        for primary in primaries:
            for gift in gifts:
                if primary.article == gift.article:
                    same_product_skipped = True
                    continue
                pair_key = (primary.article, gift.article, row.action)
                if pair_key in seen_pairs:
                    continue
                seen_pairs.add(pair_key)

                primary_disp = (
                    (primary.article_for_display or primary.article)
                    if multi_expanded
                    else row.primary_display
                )
                gift_disp = (
                    (gift.article_for_display or gift.article)
                    if multi_expanded
                    else row.gift_display
                )
                existing_display = _format_gift_list(primary.gifts, catalog)
                row_plans.append(
                    GiftPlan(
                        primary_display=primary_disp,
                        gift_display=gift_disp,
                        primary_article=primary.article,
                        gift_article=gift.article,
                        row_number=row.row_number,
                        action=row.action,
                        error="",
                        user_primary=row.primary_display,
                        user_gift=row.gift_display,
                        existing_gifts=existing_display,
                        updated_gifts=existing_display,
                    )
                )

        if not row_plans and same_product_skipped:
            primary = primaries[0]
            gift = gifts[0]
            plan.append(
                GiftPlan(
                    primary_display=row.primary_display,
                    gift_display=row.gift_display,
                    primary_article=primary.article,
                    gift_article=gift.article,
                    row_number=row.row_number,
                    action=row.action,
                    error="Основний товар і подарунок не можуть бути одним товаром.",
                    user_primary=row.primary_display,
                    user_gift=row.gift_display,
                )
            )
        else:
            plan.extend(row_plans)

    grouped_ready: dict[str, list[GiftPlan]] = {}
    for item in plan:
        if item.ready:
            grouped_ready.setdefault(item.primary_article, []).append(item)

    updated_by_primary: dict[str, tuple[str, ...]] = {}
    for primary_article, group in grouped_ready.items():
        product = catalog.by_article.get(primary_article)
        if product is not None:
            mutated, _ = mutate_gifts(product.gifts, group)
            updated_by_primary[primary_article] = _format_gift_list(mutated, catalog)

    return [
        replace(item, updated_gifts=updated_by_primary.get(item.primary_article, item.existing_gifts))
        if item.ready
        else item
        for item in plan
    ]


def mutate_gifts(existing: tuple[Any, ...], plans: list[GiftPlan]) -> tuple[list[Any], dict[tuple[str, str], str]]:
    gifts = list(existing)
    messages: dict[tuple[str, str], str] = {}
    for plan in plans:
        key = plan.gift_article.casefold()
        positions = [index for index, value in enumerate(gifts) if gift_article(value).casefold() == key]
        row_key = (plan.primary_article, plan.gift_article)
        if plan.action == "delete":
            if positions:
                gifts = [value for value in gifts if gift_article(value).casefold() != key]
                messages[row_key] = "Подарунок буде видалено."
            else:
                messages[row_key] = "Цей подарунок не був призначений товару."
        elif positions:
            messages[row_key] = "Подарунок уже призначений товару."
        else:
            gifts.append(plan.gift_article)
            messages[row_key] = "Подарунок буде додано."
    return gifts, messages


class HoroshopClient:
    def __init__(self, settings: Settings, credentials: Credentials) -> None:
        self.settings = settings
        self.credentials = credentials
        self.session = requests.Session()
        self._token = credentials.token

    def _post(self, endpoint: str, payload: dict[str, Any]) -> dict[str, Any]:
        try:
            response = self.session.post(
                endpoint_url(self.settings.domain, endpoint),
                json=payload,
                timeout=self.settings.request_timeout_seconds,
            )
            response.raise_for_status()
            data = response.json()
        except requests.RequestException as error:
            raise HoroshopGiftsError(f"Horoshop API request failed: {error}") from error
        except ValueError as error:
            raise HoroshopGiftsError("Horoshop API returned a non-JSON response.") from error
        if not isinstance(data, dict):
            raise HoroshopGiftsError("Horoshop API returned an invalid JSON response.")
        if str(data.get("status", "")).upper() in {"ERROR", "EXCEPTION"}:
            raise HoroshopGiftsError(str(data))
        return data

    def token(self) -> str:
        if self._token:
            return self._token
        login_val = self.credentials.login or self.settings.login
        password_val = self.credentials.password or self.settings.password
        if not login_val or not password_val:
            raise HoroshopGiftsError("Вкажіть логін і пароль API або чинний токен.")
        response = self._post("/api/auth/", {"login": login_val, "password": password_val})
        token = response.get("response", {}).get("token")
        if not token:
            raise HoroshopGiftsError("Хорошоп не повернув токен авторизації.")
        self._token = str(token)
        return self._token

    def export_catalog(self) -> list[dict[str, Any]]:
        offset = 0
        products: list[dict[str, Any]] = []
        while True:
            response = self._post(
                "/api/catalog/export/",
                {
                    "token": self.token(),
                    "offset": offset,
                    "limit": DEFAULT_EXPORT_LIMIT,
                    "includedParams": ["article_for_display", "gifts"],
                },
            )
            nested = response.get("response")
            page = nested.get("products") if isinstance(nested, dict) else response.get("products")
            if not isinstance(page, list):
                raise HoroshopGiftsError("Експорт каталогу не містить товарів.")
            products.extend(item for item in page if isinstance(item, dict))
            if len(page) < DEFAULT_EXPORT_LIMIT:
                return products
            offset += DEFAULT_EXPORT_LIMIT

    def import_products(self, products: list[dict[str, Any]]) -> dict[str, Any]:
        if not products:
            return {"status": "OK", "response": {"log": []}}
        return self._post("/api/catalog/import/", {"token": self.token(), "products": products})


def import_results(response: dict[str, Any]) -> dict[str, tuple[bool, str]]:
    status = str(response.get("status", "")).upper()
    nested = response.get("response")
    log_items = nested.get("log", []) if isinstance(nested, dict) else []
    results: dict[str, tuple[bool, str]] = {}
    if not isinstance(log_items, list):
        return results
    for entry in log_items:
        if not isinstance(entry, dict):
            continue
        article = normalize(entry.get("article"))
        info = entry.get("info", [])
        codes: list[Any] = []
        messages: list[str] = []
        if isinstance(info, list):
            for item in info:
                if isinstance(item, dict):
                    codes.append(item.get("code"))
                    messages.append(normalize(item.get("message")))
        results[article] = (0 in codes or (status == "OK" and not codes), "; ".join(message for message in messages if message))
    return results

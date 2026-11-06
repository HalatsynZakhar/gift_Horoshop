from __future__ import annotations

import asyncio
import base64
import datetime
import hashlib
import hmac
import logging
import sys
import threading
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, Response

from horoshop_gifts import (
    ArticleMappings,
    CatalogIndex,
    Credentials,
    GiftAssociation,
    GiftPlan,
    GiftRow,
    HoroshopClient,
    HoroshopGiftsError,
    Settings,
    build_excel_template,
    build_failed_excel,
    build_mappings_excel,
    build_registry_excel,
    build_report_excel,
    import_results,
    load_article_mappings,
    load_settings,
    mutate_gifts,
    normalize,
    parse_article_mappings,
    parse_excel_gifts,
    prepare_plan,
    save_article_mappings,
)
from images_feed import images_feed


PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_DIR / "config.json"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
MAPPINGS_FILE = PROJECT_DIR / "data" / "article_mappings.json"

SESSION_COOKIE_NAME = "gifts_session"
ALT_COOKIE_NAME = "admin_session"
SESSION_SALT = "gifts_horoshop_auth_salt_v1"


def get_session_token(login: str, password: str) -> str:
    key = f"{login}:{password}:{SESSION_SALT}".encode("utf-8")
    return hashlib.sha256(key).hexdigest()


def is_request_authenticated(request: Request, runtime_settings: Settings) -> bool:
    if not runtime_settings.auth_enabled:
        return True

    expected_token = get_session_token(runtime_settings.auth_login, runtime_settings.auth_password)
    cookie_token = request.cookies.get(SESSION_COOKIE_NAME) or request.cookies.get(ALT_COOKIE_NAME)
    if cookie_token and hmac.compare_digest(cookie_token, expected_token):
        return True

    auth_header = request.headers.get("Authorization", "")
    if auth_header.startswith("Basic "):
        try:
            encoded = auth_header.split(" ", 1)[1].strip()
            decoded = base64.b64decode(encoded).decode("utf-8")
            username, password = decoded.split(":", 1)
            if hmac.compare_digest(username, runtime_settings.auth_login) and hmac.compare_digest(password, runtime_settings.auth_password):
                return True
        except Exception:
            pass

    return False

settings: Settings | None = None
logger = logging.getLogger(__name__)
service_output_stream: "PublicLogStream | None" = None
mapping_lock = threading.Lock()


class PublicLogStream:
    def __init__(self, path: Path, fallback_path: Path) -> None:
        self.path = path
        self.fallback_path = fallback_path
        self.encoding = "utf-8"

    def write(self, message: str) -> int:
        if not message:
            return 0
        for path in (self.path, self.fallback_path):
            try:
                path.parent.mkdir(parents=True, exist_ok=True)
                with path.open("a", encoding=self.encoding) as file:
                    file.write(message)
                break
            except OSError:
                continue
        return len(message)

    def flush(self) -> None:
        return None

    def isatty(self) -> bool:
        return False


def get_settings() -> Settings:
    global settings
    if settings is None:
        if not CONFIG_FILE.exists():
            raise RuntimeError(
                f"Configuration file was not found: {CONFIG_FILE}. "
                "Create config.json from config.example.json and enter the actual settings."
            )
        settings = load_settings(CONFIG_FILE)
    return settings


def configure_service_output(runtime_settings: Settings) -> None:
    global service_output_stream
    fallback_path = PROJECT_DIR / "logs" / "horoshop_gifts.log"
    selected_path = runtime_settings.public_log_file
    try:
        selected_path.parent.mkdir(parents=True, exist_ok=True)
        selected_path.touch(exist_ok=True)
    except OSError:
        selected_path = fallback_path
        selected_path.parent.mkdir(parents=True, exist_ok=True)
        selected_path.touch(exist_ok=True)
    service_output_stream = PublicLogStream(selected_path, fallback_path)
    sys.stdout = service_output_stream
    sys.stderr = service_output_stream
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        stream=service_output_stream,
        force=True,
    )
    logger.info("Service output is writing to %s", selected_path)


def credentials_from_json(data: dict[str, Any]) -> Credentials:
    runtime_settings = get_settings()
    login = normalize(data.get("login")) or runtime_settings.login
    password = normalize(data.get("password")) or runtime_settings.password
    token = normalize(data.get("token"))
    if not token and (not login or not password):
        raise HoroshopGiftsError("Вкажіть логін і пароль API або чинний токен.")
    return Credentials(login=login, password=password, token=token)


def parse_bool(value: Any, default: bool = False) -> bool:
    text = normalize(value).casefold()
    if not text:
        return default
    return text not in {"0", "false", "ні", "нет", "no", "off"}


def stored_mappings() -> ArticleMappings:
    with mapping_lock:
        return load_article_mappings(MAPPINGS_FILE)


def add_to_stored_mappings(uploaded_mappings: ArticleMappings) -> ArticleMappings:
    with mapping_lock:
        combined = load_article_mappings(MAPPINGS_FILE).merged_with(uploaded_mappings)
        save_article_mappings(MAPPINGS_FILE, combined)
        return combined


async def uploaded_mappings(form: Any) -> ArticleMappings | None:
    mapping_file = form.get("mapping_file")
    if mapping_file is None or not hasattr(mapping_file, "read") or not normalize(getattr(mapping_file, "filename", "")):
        return None
    if not str(getattr(mapping_file, "filename", "")).lower().endswith((".xlsx", ".xlsm")):
        raise HoroshopGiftsError("Файл сопоставлень має бути .xlsx або .xlsm.")
    mapping_contents = await mapping_file.read()
    if not mapping_contents or len(mapping_contents) > MAX_UPLOAD_BYTES:
        raise HoroshopGiftsError("Файл сопоставлень порожній або перевищує 20 МБ.")
    return await asyncio.to_thread(parse_article_mappings, mapping_contents)


async def upload_bytes(request: Request) -> tuple[bytes, Credentials, bool, ArticleMappings]:
    form = await request.form()
    uploaded = form.get("file")
    if uploaded is None or not hasattr(uploaded, "read"):
        raise HoroshopGiftsError("Оберіть Excel-файл .xlsx або .xlsm.")
    filename = str(getattr(uploaded, "filename", "")).lower()
    if not filename.endswith((".xlsx", ".xlsm")):
        raise HoroshopGiftsError("Підтримуються лише Excel-файли .xlsx та .xlsm.")
    contents = await uploaded.read()
    if not contents:
        raise HoroshopGiftsError("Завантажений Excel-файл порожній.")
    if len(contents) > MAX_UPLOAD_BYTES:
        raise HoroshopGiftsError("Excel-файл перевищує обмеження 20 МБ.")
    mappings = stored_mappings()
    additional_mappings = await uploaded_mappings(form)
    if additional_mappings is not None:
        mappings = await asyncio.to_thread(add_to_stored_mappings, additional_mappings)
    normalize_articles = parse_bool(form.get("normalize_articles"), True)
    return contents, credentials_from_json(dict(form)), normalize_articles, mappings


def serialise_plan(item: GiftPlan, status: str | None = None, message: str | None = None) -> dict[str, Any]:
    primary_key = item.primary_display or item.user_primary or item.primary_article
    gift_key = item.gift_display or item.user_gift or item.gift_article
    return {
        "primary_display": item.primary_display,
        "gift_display": item.gift_display,
        "user_primary": item.user_primary or item.primary_display,
        "user_gift": item.user_gift or item.gift_display,
        "primary_article": item.primary_article,
        "gift_article": item.gift_article,
        "primary_info": images_feed.get_product_meta(primary_key),
        "gift_info": images_feed.get_product_meta(gift_key),
        "row_number": item.row_number,
        "action": item.action,
        "status": status or ("ready" if item.ready else "error"),
        "message": item.error if message is None else message,
    }


def with_failed_excel(result: dict[str, Any], failures: list[tuple[GiftPlan, str, str]]) -> dict[str, Any]:
    if failures:
        result["failed_excel"] = base64.b64encode(build_failed_excel(failures)).decode("ascii")
        result["failed_excel_name"] = "horoshop_gifts_failed.xlsx"
    return result


def with_report_excel(
    result: dict[str, Any],
    entries: list[tuple[GiftPlan, str, str]],
    filename: str = "horoshop_gifts_report.xlsx",
) -> dict[str, Any]:
    if entries:
        result["report_excel"] = base64.b64encode(build_report_excel(entries)).decode("ascii")
        result["report_excel_name"] = filename
    return result


def catalog_for(credentials: Credentials) -> tuple[CatalogIndex, HoroshopClient]:
    client = HoroshopClient(get_settings(), credentials)
    return CatalogIndex.from_raw(client.export_catalog()), client


def preview_rows(
    rows: list[GiftRow],
    credentials: Credentials,
    normalize_articles: bool = True,
    mappings: ArticleMappings | None = None,
) -> dict[str, Any]:
    catalog, _ = catalog_for(credentials)
    plan = prepare_plan(rows, catalog, normalize_articles=normalize_articles, mappings=mappings)
    grouped: dict[str, list[GiftPlan]] = {}
    for item in plan:
        if item.ready:
            grouped.setdefault(item.primary_article, []).append(item)
    pending_messages: dict[tuple[str, str], str] = {}
    for primary_article, group in grouped.items():
        product = catalog.by_article.get(primary_article)
        if product is not None:
            _, messages = mutate_gifts(product.gifts, group)
            pending_messages.update(messages)

    items: list[dict[str, Any]] = []
    failures: list[tuple[GiftPlan, str, str]] = []
    report_entries: list[tuple[GiftPlan, str, str]] = []
    for item in plan:
        if item.ready:
            msg = pending_messages.get((item.primary_article, item.gift_article), "")
            serialized = serialise_plan(item, "ready", msg)
            items.append(serialized)
            report_entries.append((item, "Готово до операції", msg))
        else:
            serialized = serialise_plan(item, "error")
            items.append(serialized)
            failures.append((item, "Помилка", serialized["message"]))

    result: dict[str, Any] = {
        "items": items,
        "ready": sum(item.ready for item in plan),
        "errors": sum(not item.ready for item in plan),
    }
    result = with_failed_excel(result, failures)
    result = with_report_excel(result, report_entries, "horoshop_gifts_preview.xlsx")
    return result


def execute_rows(
    rows: list[GiftRow],
    credentials: Credentials,
    normalize_articles: bool = True,
    mappings: ArticleMappings | None = None,
) -> dict[str, Any]:
    runtime_settings = get_settings()
    catalog, client = catalog_for(credentials)
    plan = prepare_plan(rows, catalog, normalize_articles=normalize_articles, mappings=mappings)
    grouped: dict[str, list[GiftPlan]] = {}
    for item in plan:
        if item.ready:
            grouped.setdefault(item.primary_article, []).append(item)

    payload: list[dict[str, Any]] = []
    pending_messages: dict[tuple[str, str], str] = {}
    for primary_article, group in grouped.items():
        product = catalog.by_article.get(primary_article)
        if product is None:
            continue
        gifts, messages = mutate_gifts(product.gifts, group)
        pending_messages.update(messages)
        payload.append({"article": primary_article, "gifts": gifts})

    api_results: dict[str, tuple[bool, str]] = {}
    for start in range(0, len(payload), runtime_settings.batch_size):
        api_results.update(import_results(client.import_products(payload[start : start + runtime_settings.batch_size])))

    result_items: list[dict[str, Any]] = []
    failures: list[tuple[GiftPlan, str, str]] = []
    report_entries: list[tuple[GiftPlan, str, str]] = []
    for item in plan:
        if not item.ready:
            serialized = serialise_plan(item, "invalid")
            result_items.append(serialized)
            failures.append((item, "Помилка перевірки", serialized["message"]))
            continue
        success, api_message = api_results.get(
            item.primary_article, (False, "API не повернуло результат для основного товару.")
        )
        local_message = pending_messages.get((item.primary_article, item.gift_article), "")
        message = "; ".join(value for value in (local_message, api_message) if value)
        serialized = serialise_plan(item, "synced" if success else "error", message)
        result_items.append(serialized)
        if success:
            report_entries.append((item, "Виконано", message))
        else:
            failures.append((item, "Помилка API", message))

    result: dict[str, Any] = {
        "items": result_items,
        "imported": sum(item["status"] == "synced" for item in result_items),
        "errors": sum(item["status"] in {"error", "invalid"} for item in result_items),
    }
    result = with_failed_excel(result, failures)
    result = with_report_excel(result, report_entries, "horoshop_gifts_report.xlsx")
    return result


def list_associations(credentials: Credentials) -> dict[str, Any]:
    catalog, _ = catalog_for(credentials)
    associations = catalog.associations()
    return {
        "items": [
            {
                "primary_display": item.primary_display,
                "gift_display": item.gift_display,
                "primary_article": item.primary_article,
                "gift_article": item.gift_article,
                "primary_info": images_feed.get_product_meta(item.primary_display or item.primary_article),
                "gift_info": images_feed.get_product_meta(item.gift_display or item.gift_article),
            }
            for item in associations
        ],
        "count": len(associations),
        "images_loaded": images_feed.is_loaded(),
    }


def rows_from_pairs(raw_pairs: Any, action: str) -> list[GiftRow]:
    if not isinstance(raw_pairs, list):
        raise HoroshopGiftsError("Передайте список зв'язків товарів і подарунків.")
    rows: list[GiftRow] = []
    for index, pair in enumerate(raw_pairs, start=1):
        if not isinstance(pair, dict):
            raise HoroshopGiftsError("Кожен зв'язок має бути об'єктом.")
        primary = normalize(pair.get("primary_display"))
        gift = normalize(pair.get("gift_display"))
        if not primary or not gift:
            raise HoroshopGiftsError("Вкажіть артикул основного товару та подарунка.")
        rows.append(GiftRow(primary, gift, index, action))
    if not rows:
        raise HoroshopGiftsError("Виберіть хоча б один зв'язок.")
    return rows


def page_html() -> str:
    return (PROJECT_DIR / "web_ui.html").read_text(encoding="utf-8")


def login_html() -> str:
    return (PROJECT_DIR / "login.html").read_text(encoding="utf-8")


app = FastAPI(title="Подарунки Хорошоп")


@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    runtime_settings = get_settings()
    path = request.url.path

    if path in {"/login", "/api/login", "/favicon.ico", "/api/auth/status"}:
        return await call_next(request)

    if not is_request_authenticated(request, runtime_settings):
        if path.startswith("/api/"):
            return JSONResponse(
                status_code=401,
                content={"detail": "Потрібна авторизація (Необхідно увійти у систему)."},
            )
        return RedirectResponse(url=f"/login?next={path}", status_code=303)

    return await call_next(request)


@app.exception_handler(Exception)
async def unexpected_error(_: Request, error: Exception) -> JSONResponse:
    logger.exception("Unhandled web request error", exc_info=error)
    return JSONResponse(status_code=500, content={"detail": "Внутрішня помилка сервера. Перевірте публічний лог."})


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request) -> Response:
    runtime_settings = get_settings()
    if is_request_authenticated(request, runtime_settings):
        return RedirectResponse(url="/", status_code=303)
    return HTMLResponse(
        content=login_html(),
        headers={"Cache-Control": "no-store, no-cache, must-revalidate, max-age=0"},
    )


@app.post("/api/login")
async def login_endpoint(request: Request) -> Response:
    runtime_settings = get_settings()
    login_val = ""
    password_val = ""

    content_type = request.headers.get("content-type", "")
    if "application/json" in content_type:
        try:
            body = await request.json()
            if isinstance(body, dict):
                login_val = normalize(body.get("login") or body.get("username"))
                password_val = normalize(body.get("password"))
        except Exception:
            pass
    else:
        try:
            form = await request.form()
            login_val = normalize(form.get("login") or form.get("username"))
            password_val = normalize(form.get("password"))
        except Exception:
            pass

    expected_login = runtime_settings.auth_login
    expected_password = runtime_settings.auth_password

    if not (hmac.compare_digest(login_val, expected_login) and hmac.compare_digest(password_val, expected_password)):
        return JSONResponse(
            status_code=401,
            content={"success": False, "detail": "Невірний логін або пароль."},
        )

    token = get_session_token(expected_login, expected_password)
    response = JSONResponse(
        content={"success": True, "message": "Авторизація успішна.", "redirect": "/"},
    )
    response.set_cookie(
        key=SESSION_COOKIE_NAME,
        value=token,
        max_age=30 * 24 * 3600,
        httponly=True,
        samesite="lax",
        secure=False,
        path="/",
    )
    response.set_cookie(
        key=ALT_COOKIE_NAME,
        value=token,
        max_age=30 * 24 * 3600,
        httponly=True,
        samesite="lax",
        secure=False,
        path="/",
    )
    return response


@app.post("/api/logout")
def logout_endpoint() -> Response:
    response = JSONResponse(content={"success": True, "redirect": "/login"})
    response.delete_cookie(key=SESSION_COOKIE_NAME, path="/")
    response.delete_cookie(key=ALT_COOKIE_NAME, path="/")
    return response


@app.get("/api/auth/status")
def auth_status(request: Request) -> dict[str, Any]:
    runtime_settings = get_settings()
    authenticated = is_request_authenticated(request, runtime_settings)
    return {
        "authenticated": authenticated,
        "auth_enabled": runtime_settings.auth_enabled,
        "user": runtime_settings.auth_login if authenticated else None,
        "has_horoshop_credentials": bool(runtime_settings.login and runtime_settings.password),
        "horoshop_login": runtime_settings.login if authenticated else None,
        "domain": runtime_settings.domain if authenticated else None,
    }


@app.get("/", response_class=HTMLResponse)
def index(request: Request) -> Response:
    runtime_settings = get_settings()
    if not is_request_authenticated(request, runtime_settings):
        return RedirectResponse(url="/login", status_code=303)
    return HTMLResponse(
        content=page_html(),
        headers={"Cache-Control": "no-store, no-cache, must-revalidate, max-age=0"},
    )


@app.get("/api/template")
def download_template() -> Response:
    return Response(
        content=build_excel_template(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="horoshop_gifts_template.xlsx"',
            "Cache-Control": "no-store",
        },
    )


@app.get("/api/mappings/info")
def mappings_info() -> dict[str, int]:
    mappings = stored_mappings()
    return {"sources": len(mappings.entries), "rules": sum(len(targets) for targets in mappings.entries.values())}


@app.get("/api/mappings/export")
def export_mappings() -> Response:
    content = build_mappings_excel(stored_mappings())
    return Response(
        content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="horoshop_article_mappings.xlsx"',
            "Cache-Control": "no-store",
        },
    )


@app.post("/api/mappings")
async def upload_mappings_endpoint(request: Request) -> dict[str, int]:
    try:
        mappings = await uploaded_mappings(await request.form())
        if mappings is None:
            raise HoroshopGiftsError("Оберіть Excel-файл сопоставлень.")
        stored = await asyncio.to_thread(add_to_stored_mappings, mappings)
        return {"sources": len(stored.entries), "rules": sum(len(targets) for targets in stored.entries.values())}
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/mappings/clear")
async def clear_mappings() -> dict[str, Any]:
    with mapping_lock:
        backup_name = ""
        if MAPPINGS_FILE.exists():
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = MAPPINGS_FILE.with_name(f"article_mappings_{timestamp}.json.bak")
            latest_bak = MAPPINGS_FILE.with_suffix(".json.bak")
            try:
                content = MAPPINGS_FILE.read_text(encoding="utf-8")
                backup_file.write_text(content, encoding="utf-8")
                latest_bak.write_text(content, encoding="utf-8")
                backup_name = backup_file.name
            except OSError as error:
                logger.warning("Could not backup mappings file: %s", error)
        save_article_mappings(MAPPINGS_FILE, ArticleMappings.empty())
        return {"status": "cleared", "backup": backup_name}


@app.post("/api/server/restart")
async def restart_service() -> dict[str, str]:
    def do_exit() -> None:
        import os
        import time

        time.sleep(0.5)
        os._exit(0)

    threading.Thread(target=do_exit, daemon=True).start()
    return {"status": "restarting"}


@app.post("/api/gifts/list")
async def gifts_list(request: Request) -> dict[str, Any]:
    try:
        data = await request.json()
        return await asyncio.to_thread(list_associations, credentials_from_json(data))
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/gifts/export")
async def export_gifts(request: Request) -> Response:
    try:
        data = await request.json()
        associations = await asyncio.to_thread(list_associations, credentials_from_json(data))
        contents = build_registry_excel([GiftAssociation(**item) for item in associations["items"]])
        return Response(
            content=contents,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": 'attachment; filename="horoshop_gifts_registry.xlsx"',
                "Cache-Control": "no-store",
            },
        )
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/gifts/create")
async def create_gift(request: Request) -> dict[str, Any]:
    try:
        data = await request.json()
        credentials = credentials_from_json(data)
        row = GiftRow(normalize(data.get("primary_display")), normalize(data.get("gift_display")), 1)
        if not row.primary_display or not row.gift_display:
            raise HoroshopGiftsError("Вкажіть артикул основного товару та подарунка.")
        normalize_articles = parse_bool(data.get("normalize_articles"), True)
        return await asyncio.to_thread(execute_rows, [row], credentials, normalize_articles, stored_mappings())
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/gifts/delete")
async def delete_gifts(request: Request) -> dict[str, Any]:
    try:
        data = await request.json()
        credentials = credentials_from_json(data)
        normalize_articles = parse_bool(data.get("normalize_articles"), True)
        rows = rows_from_pairs(data.get("pairs"), "delete")
        use_mappings = parse_bool(data.get("use_mappings"), False)
        mappings = stored_mappings() if use_mappings else ArticleMappings.empty()
        return await asyncio.to_thread(execute_rows, rows, credentials, normalize_articles, mappings)
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/preview")
async def preview(request: Request) -> dict[str, Any]:
    try:
        contents, credentials, normalize_articles, mappings = await upload_bytes(request)
        return await asyncio.to_thread(
            preview_rows,
            parse_excel_gifts(contents),
            credentials,
            normalize_articles,
            mappings,
        )
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/import")
async def import_gifts(request: Request) -> dict[str, Any]:
    try:
        contents, credentials, normalize_articles, mappings = await upload_bytes(request)
        return await asyncio.to_thread(
            execute_rows,
            parse_excel_gifts(contents),
            credentials,
            normalize_articles,
            mappings,
        )
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.get("/api/products/info")
def products_info(articles: str = "") -> dict[str, Any]:
    raw_list = [a.strip() for a in articles.replace(";", ",").split(",") if a.strip()]
    return {
        "items": [images_feed.get_product_meta(art) for art in raw_list],
        "images_loaded": images_feed.is_loaded(),
    }


@app.get("/api/images/status")
def images_status() -> dict[str, Any]:
    return {
        "loaded": images_feed.is_loaded(),
        "count": images_feed.count(),
        "url": images_feed.xml_url,
        "error": images_feed.last_error(),
    }


@app.post("/api/images/refresh")
def images_refresh() -> dict[str, Any]:
    threading.Thread(target=images_feed.load, kwargs={"force_refresh": True}, daemon=True).start()
    return {"status": "refreshing", "url": images_feed.xml_url}


def run_server() -> None:
    import uvicorn

    runtime_settings = get_settings()
    configure_service_output(runtime_settings)
    threading.Thread(target=images_feed.load, daemon=True).start()
    uvicorn.run(app, host=runtime_settings.host, port=runtime_settings.port)


# Фонове завантаження XML-фіду зображень при запуску
threading.Thread(target=images_feed.load, daemon=True).start()


if __name__ == "__main__":
    run_server()

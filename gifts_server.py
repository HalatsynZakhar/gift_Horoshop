from __future__ import annotations

import asyncio
import logging
import sys
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

from horoshop_gifts import (
    CatalogIndex,
    Credentials,
    GiftPlan,
    GiftRow,
    HoroshopClient,
    HoroshopGiftsError,
    Settings,
    build_excel_template,
    build_registry_excel,
    import_results,
    load_settings,
    normalize,
    parse_excel_gifts,
    prepare_plan,
    mutate_gifts,
)


PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_DIR / "config.json"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024

settings: Settings | None = None
logger = logging.getLogger(__name__)
service_output_stream: "PublicLogStream | None" = None


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
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s", stream=service_output_stream, force=True)
    logger.info("Service output is writing to %s", selected_path)


def credentials_from_json(data: dict[str, Any]) -> Credentials:
    credentials = Credentials(
        login=normalize(data.get("login")),
        password=normalize(data.get("password")),
        token=normalize(data.get("token")),
    )
    if not credentials.token and (not credentials.login or not credentials.password):
        raise HoroshopGiftsError("Вкажіть логін і пароль API або чинний токен.")
    return credentials


async def upload_bytes(request: Request) -> tuple[bytes, Credentials]:
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
    return contents, credentials_from_json(dict(form))


def serialise_plan(item: GiftPlan, status: str | None = None, message: str | None = None) -> dict[str, Any]:
    return {
        "primary_display": item.primary_display,
        "gift_display": item.gift_display,
        "primary_article": item.primary_article,
        "gift_article": item.gift_article,
        "row_number": item.row_number,
        "action": item.action,
        "status": status or ("ready" if item.ready else "error"),
        "message": item.error if message is None else message,
    }


def catalog_for(credentials: Credentials) -> tuple[CatalogIndex, HoroshopClient]:
    client = HoroshopClient(get_settings(), credentials)
    return CatalogIndex.from_raw(client.export_catalog()), client


def preview_rows(rows: list[GiftRow], credentials: Credentials) -> dict[str, Any]:
    catalog, _ = catalog_for(credentials)
    plan = prepare_plan(rows, catalog)
    return {"items": [serialise_plan(item) for item in plan], "ready": sum(item.ready for item in plan), "errors": sum(not item.ready for item in plan)}


def execute_rows(rows: list[GiftRow], credentials: Credentials) -> dict[str, Any]:
    runtime_settings = get_settings()
    catalog, client = catalog_for(credentials)
    plan = prepare_plan(rows, catalog)
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
    for item in plan:
        if not item.ready:
            result_items.append(serialise_plan(item, "invalid"))
            continue
        success, api_message = api_results.get(item.primary_article, (False, "API не повернуло результат для основного товару."))
        local_message = pending_messages.get((item.primary_article, item.gift_article), "")
        message = "; ".join(value for value in (local_message, api_message) if value)
        result_items.append(serialise_plan(item, "synced" if success else "error", message))
    return {
        "items": result_items,
        "imported": sum(item["status"] == "synced" for item in result_items),
        "errors": sum(item["status"] in {"error", "invalid"} for item in result_items),
    }


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
            }
            for item in associations
        ],
        "count": len(associations),
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


app = FastAPI(title="Подарунки Хорошоп")


@app.exception_handler(Exception)
async def unexpected_error(_: Request, error: Exception) -> JSONResponse:
    logger.exception("Unhandled web request error", exc_info=error)
    return JSONResponse(status_code=500, content={"detail": "Внутрішня помилка сервера. Перевірте публічний лог."})


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return page_html()


@app.get("/api/template")
def download_template() -> Response:
    return Response(content=build_excel_template(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": 'attachment; filename="horoshop_gifts_template.xlsx"', "Cache-Control": "no-store"})


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
        from horoshop_gifts import GiftAssociation
        contents = build_registry_excel([GiftAssociation(**item) for item in associations["items"]])
        return Response(content=contents, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": 'attachment; filename="horoshop_gifts_registry.xlsx"', "Cache-Control": "no-store"})
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
        return await asyncio.to_thread(execute_rows, [row], credentials)
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/gifts/delete")
async def delete_gifts(request: Request) -> dict[str, Any]:
    try:
        data = await request.json()
        credentials = credentials_from_json(data)
        return await asyncio.to_thread(execute_rows, rows_from_pairs(data.get("pairs"), "delete"), credentials)
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/preview")
async def preview(request: Request) -> dict[str, Any]:
    try:
        contents, credentials = await upload_bytes(request)
        return await asyncio.to_thread(preview_rows, parse_excel_gifts(contents), credentials)
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/import")
async def import_gifts(request: Request) -> dict[str, Any]:
    try:
        contents, credentials = await upload_bytes(request)
        return await asyncio.to_thread(execute_rows, parse_excel_gifts(contents), credentials)
    except (HoroshopGiftsError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


def run_server() -> None:
    import uvicorn

    runtime_settings = get_settings()
    configure_service_output(runtime_settings)
    uvicorn.run(app, host=runtime_settings.host, port=runtime_settings.port)


if __name__ == "__main__":
    run_server()

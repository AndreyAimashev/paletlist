# -*- coding: utf-8 -*-
"""Упаковочные листы для обычных клиентов (не «Арнест Юнирусь»): PDF/HTML через WeasyPrint.

Сейчас выводится пустая страница A4 в альбомной ориентации на каждую паллету в сборке — заготовка под новый макет.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
STYLES_PATH = BASE_DIR / "packing_sheet_generic_styles.css"


def _is_arnest_unirus_client(client: str) -> bool:
    n = str(client or "").strip().replace("  ", " ").lower()
    return n == "арнест юнирусь"


def _migrate_pallet(p: dict[str, Any]) -> dict[str, Any]:
    """Нормализация id и номера паллеты (как раньше, для совместимости со сборкой)."""
    pid = int(p.get("id") or 0)
    pallet_number = str(p.get("palletNumber") or "").strip()
    if isinstance(p.get("slots"), list):
        return {"id": pid, "palletNumber": pallet_number, "slots": p["slots"]}
    return {"id": pid, "palletNumber": pallet_number, "slots": []}


def _load_styles() -> str:
    if STYLES_PATH.is_file():
        return STYLES_PATH.read_text(encoding="utf-8")
    return (
        "@page{size:A4 landscape;margin:10mm}"
        "body{margin:0}"
        ".generic-pallet-sheet{page-break-after:always;min-height:180mm;width:100%;box-sizing:border-box}"
        ".generic-pallet-sheet:last-of-type{page-break-after:auto}"
    )


def build_generic_packing_sheets_html(detail: dict[str, Any]) -> tuple[str | None, str | None]:
    """Полный HTML: по одной пустой альбомной странице на паллету."""
    if _is_arnest_unirus_client(str(detail.get("client") or "")):
        return None, "arnest_client"
    st = detail.get("assemble_state")
    if not isinstance(st, dict):
        return None, "no_assembly"
    pallets_raw = st.get("pallets")
    if not isinstance(pallets_raw, list) or len(pallets_raw) == 0:
        return None, "no_pallets"
    pallets = [_migrate_pallet(p) for p in pallets_raw if isinstance(p, dict)]
    if not pallets:
        return None, "no_pallets"

    sections: list[str] = []
    for _pal in pallets:
        sections.append('<div class="generic-pallet-sheet" aria-hidden="true"></div>')

    styles = _load_styles()
    doc = (
        '<!DOCTYPE html>\n<html lang="ru"><head><meta charset="utf-8"/>'
        '<meta name="viewport" content="width=device-width, initial-scale=1"/>'
        "<title>Упаковочные листы</title><style>"
        + styles
        + "</style></head><body>\n"
        + "\n".join(sections)
        + "\n</body></html>"
    )
    return doc, None


def build_generic_packing_sheets_pdf_bytes(detail: dict[str, Any]) -> tuple[bytes | None, str | None]:
    """Пустой альбомный PDF на основе того же HTML (WeasyPrint)."""
    html, err = build_generic_packing_sheets_html(detail)
    if err:
        return None, err
    try:
        from weasyprint import HTML as WeasyHTML  # noqa: PLC0415
    except (ImportError, OSError):
        return None, "no_pdf_engine"
    try:
        base_url = BASE_DIR.as_uri() + "/"
        pdf = WeasyHTML(string=html, base_url=base_url).write_pdf()
    except Exception as exc:  # noqa: BLE001
        return None, "pdf_render:" + str(exc)
    if not pdf:
        return None, "pdf_empty"
    return bytes(pdf), None


def generic_packing_error_message(code: str) -> str:
    if code.startswith("pdf_render:"):
        return "Не удалось сформировать PDF: " + code.partition(":")[2].strip()
    return {
        "arnest_client": "Для клиента «Арнест Юнирусь» используйте PDF со штрих-кодами.",
        "no_assembly": "Нет данных сборки (assemble_state). Сохраните сборку в разделе «Сборка заказа».",
        "no_pallets": "Нет паллет в сборке.",
        "no_pdf_engine": (
            "Не удалось загрузить WeasyPrint (HTML→PDF). На Linux установите системные "
            "библиотеки Pango/Cairo (см. документацию WeasyPrint) и пакет weasyprint из requirements.txt."
        ),
        "pdf_empty": "Сформирован пустой PDF.",
    }.get(code, str(code))

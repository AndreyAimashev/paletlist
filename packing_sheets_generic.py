# -*- coding: utf-8 -*-
"""Упаковочные листы для обычных клиентов (не «Арнест Юнирусь»): PDF/HTML через WeasyPrint."""

from __future__ import annotations

import html
import re
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
STYLES_PATH = BASE_DIR / "packing_sheet_generic_styles.css"


def _is_arnest_unirus_client(client: str) -> bool:
    n = str(client or "").strip().replace("  ", " ").lower()
    return n == "арнест юнирусь"


def _ship_date_ru(ship_date: str) -> str:
    s = str(ship_date or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        y, mo, d = m.groups()
        return f"{d}.{mo}.{y}"
    return s or "—"


def _packing_list_number(order_id: Any, pallet_number_raw: str, pallet_index: int) -> str:
    """Номер упаковочного листа: номер заказа и номер паллеты подряд (без разделителя)."""
    oid = str(order_id if order_id is not None else "").strip() or "0"
    raw = str(pallet_number_raw or "").strip()
    part = re.sub(r"\s+", "", raw) if raw else str(pallet_index)
    return f"{oid}{part}"


def _migrate_pallet(p: dict[str, Any]) -> dict[str, Any]:
    """Нормализация id и номера паллеты (как раньше, для совместимости со сборкой)."""
    pid = int(p.get("id") or 0)
    pallet_number = str(p.get("palletNumber") or "").strip()
    if isinstance(p.get("slots"), list):
        return {"id": pid, "palletNumber": pallet_number, "slots": p["slots"]}
    return {"id": pid, "palletNumber": pallet_number, "slots": []}


def _pallet_no_display(pnum_raw: str, pallet_index: int) -> str:
    """Номер паллеты для строки 2: цифра в рамке; если пусто — порядковый номер."""
    raw = str(pnum_raw or "").strip()
    if not raw:
        return str(pallet_index)
    try:
        return str(int(float(raw.replace(",", "."))))
    except ValueError:
        return raw


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
    """Полный HTML: альбомная страница на паллету (строки 1–2 шапки)."""
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

    order_id = detail.get("id")
    ship_ru = _ship_date_ru(str(detail.get("ship_date") or ""))
    ship_e = html.escape(ship_ru, quote=True)
    total_pallets = len(pallets)
    total_e = html.escape(str(total_pallets), quote=True)

    sections: list[str] = []
    for idx, pal in enumerate(pallets, start=1):
        pnum_raw = str(pal.get("palletNumber") or "").strip()
        list_no = _packing_list_number(order_id, pnum_raw, idx)
        list_no_e = html.escape(list_no, quote=True)
        pallet_disp = _pallet_no_display(pnum_raw, idx)
        pallet_disp_e = html.escape(pallet_disp, quote=True)

        row1 = (
            '<div class="generic-row1">'
            '<span class="generic-r1-label">Упаковочный лист №</span>'
            '<span class="generic-r1-gap"> </span>'
            f'<span class="generic-r1-frame">{list_no_e}</span>\t'
            '<span class="generic-r1-label">Дата</span>'
            '<span class="generic-r1-gap"> </span>'
            f'<span class="generic-r1-frame">{ship_e}</span>'
            "</div>"
        )
        row2 = (
            '<div class="generic-row2">'
            '<span class="generic-r2-label">Номер паллета</span>'
            '<span class="generic-r2-gap"> </span>'
            '<table class="generic-r2-triplet" role="presentation" aria-label="Номер паллеты из общего числа">'
            "<tr>"
            f'<td class="generic-r2-cell">{pallet_disp_e}</td>'
            '<td class="generic-r2-cell">из</td>'
            f'<td class="generic-r2-cell">{total_e}</td>'
            "</tr></table></div>"
        )
        sections.append(f'<div class="generic-pallet-sheet">{row1}{row2}</div>')

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
    """Альбомный PDF из того же HTML (WeasyPrint)."""
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

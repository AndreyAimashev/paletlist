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


SUPPLIER_LINE = "М.К. АСЕПТИКА ООО"


def _fmt_ru_num(n: float, *, max_decimals: int = 3) -> str:
    if not isinstance(n, (int, float)) or n != n:
        return "—"
    if abs(n - round(n)) < 1e-9:
        s = f"{int(round(n)):,}".replace(",", "\xa0")
        return s
    fmt = f"{{:.{max_decimals}f}}"
    t = fmt.format(n).rstrip("0").rstrip(".")
    return t.replace(".", ",")


def _normalize_slot(slot: dict[str, Any]) -> dict[str, Any]:
    li = slot.get("lineIndex")
    line_index: int | None = None
    if li not in ("", None):
        try:
            n = int(float(str(li).replace(",", ".")))
            if n >= 0:
                line_index = n
        except (TypeError, ValueError):
            pass
    mode = slot.get("mode")
    mode_s = "rows" if mode == "rows" else "direct"
    du = slot.get("directUnit")
    direct_unit = "piece" if du == "piece" else "box"
    return {
        "lineIndex": line_index,
        "mode": mode_s,
        "directUnit": direct_unit,
        "directQty": slot.get("directQty", ""),
        "fullRows": slot.get("fullRows", ""),
        "partialBoxes": slot.get("partialBoxes", ""),
    }


def _slot_to_alloc(slot: dict[str, Any]) -> dict[str, Any]:
    return {
        "mode": slot["mode"],
        "directUnit": slot["directUnit"],
        "directQty": slot["directQty"],
        "fullRows": slot["fullRows"],
        "partialBoxes": slot["partialBoxes"],
    }


def _allocation_to_pieces(it: dict[str, Any], alloc: dict[str, Any]) -> float:
    pib = max(0, int(it.get("pieces_in_box") or 0))
    row_lay = max(0, int(it.get("row_layout") or 0))
    if alloc.get("mode") == "rows":
        if row_lay <= 0:
            return 0.0
        try:
            fr = max(0, int(float(str(alloc.get("fullRows") or "").replace(",", "."))))
        except ValueError:
            fr = 0
        try:
            pb = max(0, int(float(str(alloc.get("partialBoxes") or "").replace(",", "."))))
        except ValueError:
            pb = 0
        boxes = fr * row_lay + pb
        return float(boxes * pib) if pib > 0 else 0.0
    try:
        dq = max(0.0, float(str(alloc.get("directQty") or "").replace(",", ".")))
    except ValueError:
        dq = 0.0
    if alloc.get("directUnit") == "box":
        return float(dq * pib) if pib > 0 else 0.0
    return dq


def _allocation_to_boxes(it: dict[str, Any], alloc: dict[str, Any]) -> float:
    pib = max(0, int(it.get("pieces_in_box") or 0))
    if pib <= 0:
        return 0.0
    pc = _allocation_to_pieces(it, alloc)
    return pc / pib


def _nomenclature_title(it: dict[str, Any]) -> str:
    name = (it.get("name") or "—").strip() or "—"
    art = (it.get("article") or "").strip()
    if art:
        return f"{art} — {name}"
    return name


def _build_pallet_lines_table_html(items: list[dict[str, Any]], pal: dict[str, Any]) -> str:
    """Таблица строки 5: агрегат по строкам заказа (lineIndex) на паллете."""
    order_keys: list[int] = []
    agg: dict[int, dict[str, float]] = {}
    for slot in pal.get("slots") or []:
        if not isinstance(slot, dict):
            continue
        ns = _normalize_slot(slot)
        li = ns["lineIndex"]
        if li is None or li < 0 or li >= len(items):
            continue
        it = items[li]
        alloc = _slot_to_alloc(ns)
        pcs = _allocation_to_pieces(it, alloc)
        box = _allocation_to_boxes(it, alloc)
        if pcs <= 1e-9 and box <= 1e-9:
            continue
        if li not in agg:
            agg[li] = {"pcs": 0.0, "boxes": 0.0}
            order_keys.append(li)
        agg[li]["pcs"] += pcs
        agg[li]["boxes"] += box

    head = (
        "<table class=\"generic-lines-table\" cellspacing=\"0\">"
        "<thead><tr>"
        '<th class="generic-th generic-th-pp">пп</th>'
        '<th class="generic-th generic-th-name">Номенклатура</th>'
        '<th class="generic-th generic-th-num">Кол-во штук</th>'
        '<th class="generic-th generic-th-num">Кол-во коробок</th>'
        '<th class="generic-th generic-th-num">Штук в коробке</th>'
        '<th class="generic-th generic-th-num">Вес (кг)</th>'
        "</tr></thead><tbody>"
    )
    if not order_keys:
        return (
            head
            + '<tr><td class="generic-td" colspan="6">Нет распределённых позиций на этой паллете.</td></tr>'
            + "</tbody></table>"
        )

    body_parts: list[str] = []
    total_boxes = 0.0
    total_weight = 0.0
    for n, li in enumerate(order_keys, start=1):
        it = items[li]
        row = agg[li]
        pcs = row["pcs"]
        box_raw = float(row["boxes"])
        # На паллете по каждому наименованию «Кол-во коробок» не меньше 1
        box = max(1.0, box_raw) if (pcs > 1e-9 or box_raw > 1e-9) else box_raw
        pib = max(0, int(it.get("pieces_in_box") or 0))
        bw = float(it.get("box_weight") or 0)
        weight = box * bw if bw > 0 and box > 0 else 0.0
        total_boxes += box
        total_weight += weight

        title = _nomenclature_title(it)
        title_e = html.escape(title, quote=True)
        pcs_s = _fmt_ru_num(pcs, max_decimals=0) if abs(pcs - round(pcs)) < 1e-6 else _fmt_ru_num(pcs, max_decimals=3)
        box_s = _fmt_ru_num(box, max_decimals=3)
        pib_s = str(pib) if pib > 0 else "—"
        w_s = _fmt_ru_num(weight, max_decimals=2) if weight > 0 else "—"

        body_parts.append(
            "<tr>"
            f'<td class="generic-td generic-td-pp">{n}</td>'
            f'<td class="generic-td generic-td-name">{title_e}</td>'
            f'<td class="generic-td generic-td-num">{html.escape(pcs_s, quote=True)}</td>'
            f'<td class="generic-td generic-td-num">{html.escape(box_s, quote=True)}</td>'
            f'<td class="generic-td generic-td-num">{html.escape(pib_s, quote=True)}</td>'
            f'<td class="generic-td generic-td-num">{html.escape(w_s, quote=True)}</td>'
            "</tr>"
        )

    box_tot_s = _fmt_ru_num(total_boxes, max_decimals=3)
    w_tot_s = _fmt_ru_num(total_weight, max_decimals=2) if total_weight > 1e-12 else "—"
    foot = (
        "<tfoot><tr class=\"generic-lines-tfoot\">"
        '<td colspan="3" class="generic-td generic-td-total-label">Итого:</td>'
        f'<td class="generic-td generic-td-num">{html.escape(box_tot_s, quote=True)}</td>'
        '<td class="generic-td generic-td-num"></td>'
        f'<td class="generic-td generic-td-num">{html.escape(w_tot_s, quote=True)}</td>'
        "</tr></tfoot>"
    )
    return head + "".join(body_parts) + "</tbody>" + foot + "</table>"


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
    """Полный HTML: альбомная страница на паллету (строки 1–5: шапка + таблица номенклатуры)."""
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

    items = detail.get("items")
    if not isinstance(items, list):
        items = []

    order_id = detail.get("id")
    ship_ru = _ship_date_ru(str(detail.get("ship_date") or ""))
    ship_e = html.escape(ship_ru, quote=True)
    total_pallets = len(pallets)
    total_e = html.escape(str(total_pallets), quote=True)
    buyer_e = html.escape(str(detail.get("client") or "").strip() or "—", quote=True)
    supplier_e = html.escape(SUPPLIER_LINE, quote=True)

    sections: list[str] = []
    for idx, pal in enumerate(pallets, start=1):
        pnum_raw = str(pal.get("palletNumber") or "").strip()
        list_no = _packing_list_number(order_id, pnum_raw, idx)
        list_no_e = html.escape(list_no, quote=True)
        pallet_disp = _pallet_no_display(pnum_raw, idx)
        pallet_disp_e = html.escape(pallet_disp, quote=True)

        row1 = (
            '<div class="generic-row1">'
            '<span class="generic-r1-left">'
            '<span class="generic-r1-label">Упаковочный лист №</span>'
            '<span class="generic-r1-gap"> </span>'
            f'<span class="generic-r1-frame">{list_no_e}</span>'
            "</span>"
            '<span class="generic-r1-right">'
            '<span class="generic-r1-label">Дата</span>'
            '<span class="generic-r1-gap"> </span>'
            f'<span class="generic-r1-frame">{ship_e}</span>'
            "</span>"
            "</div>"
        )
        row2 = (
            '<div class="generic-row2">'
            '<span class="generic-r2-label">Номер паллета</span>'
            '<span class="generic-r2-gap"> </span>'
            '<table class="generic-r2-triplet" role="presentation" aria-label="Номер паллеты из общего числа">'
            "<tr>"
            f'<td class="generic-r2-cell generic-r2-cell--side">{pallet_disp_e}</td>'
            '<td class="generic-r2-cell generic-r2-cell--mid">из</td>'
            f'<td class="generic-r2-cell generic-r2-cell--side">{total_e}</td>'
            "</tr></table></div>"
        )
        row3 = (
            '<div class="generic-row34">'
            '<span class="generic-r34-label">Поставщик:</span>'
            f'<span class="generic-r34-center">{supplier_e}</span>'
            "</div>"
        )
        row4 = (
            '<div class="generic-row34">'
            '<span class="generic-r34-label">Покупатель:</span>'
            f'<span class="generic-r34-center">{buyer_e}</span>'
            "</div>"
        )
        row5 = '<div class="generic-row5-wrap">' + _build_pallet_lines_table_html(items, pal) + "</div>"
        sections.append(f'<div class="generic-pallet-sheet">{row1}{row2}{row3}{row4}{row5}</div>')

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

# -*- coding: utf-8 -*-
"""Упаковочные листы для обычных клиентов (не «Арнест Юнирусь»): HTML по шаблону с рабочего стола.

Логика распределения по паллетам совпадает с фронтом admin (режимы direct / rows).
"""

from __future__ import annotations

import html
import re
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
STYLES_PATH = BASE_DIR / "packing_sheet_generic_styles.css"
FRAGMENT_PATH = BASE_DIR / "packing_sheet_generic_pallet_fragment.html"

# Пропорции как в образце (пиксели FastReport); 11-я колонка — хвост строки (в шаблоне 11 ячеек в ряду R0).
_COL_FRACS_11 = (4, 38, 181, 148, 75, 144, 155, 185, 56, 56, 20)


def _generic_colgroup_html() -> str:
    """Проценты ширины колонок под любую ширину страницы (WeasyPrint + landscape)."""
    s = float(sum(_COL_FRACS_11))
    parts: list[str] = []
    acc = 0.0
    for i, w in enumerate(_COL_FRACS_11):
        if i == len(_COL_FRACS_11) - 1:
            pct = round(100.0 - acc, 4)
        else:
            p = 100.0 * float(w) / s
            pct = round(p, 4)
            acc += pct
        parts.append(f'<col style="width:{pct}%" />')
    return "<colgroup>\n" + "\n".join(parts) + "\n</colgroup>"


def _is_arnest_unirus_client(client: str) -> bool:
    n = str(client or "").strip().replace("  ", " ").lower()
    return n == "арнест юнирусь"


def _supplier_default() -> str:
    import os

    return (os.environ.get("GENERIC_PACKING_SHEET_SUPPLIER") or "").strip() or "—"


def _fmt_ru_num(n: float, *, max_decimals: int = 3) -> str:
    if not isinstance(n, (int, float)) or n != n:
        return "—"
    if abs(n - round(n)) < 1e-9:
        s = f"{int(round(n)):,}".replace(",", "\xa0")
        return s
    fmt = f"{{:.{max_decimals}f}}"
    t = fmt.format(n).rstrip("0").rstrip(".")
    return t.replace(".", ",")


def _ship_date_ru(ship_date: str) -> str:
    s = str(ship_date or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        y, mo, d = m.groups()
        return f"{d}.{mo}.{y}"
    return s or "—"


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
            fr = max(0, int(str(alloc.get("fullRows") or ""), 10))
        except ValueError:
            fr = 0
        try:
            pb = max(0, int(str(alloc.get("partialBoxes") or ""), 10))
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


def _migrate_pallet(p: dict[str, Any]) -> dict[str, Any]:
    pid = int(p.get("id") or 0)
    pallet_number = str(p.get("palletNumber") or "").strip()
    if isinstance(p.get("slots"), list):
        return {"id": pid, "palletNumber": pallet_number, "slots": [_normalize_slot(s) for s in p["slots"]]}
    slots: list[dict[str, Any]] = []
    lines = p.get("lines")
    if isinstance(lines, dict):
        for k, v in lines.items():
            try:
                li = int(k)
            except (TypeError, ValueError):
                continue
            if li < 0:
                continue
            if isinstance(v, dict):
                slots.append(_normalize_slot({"lineIndex": li, **v}))
    return {"id": pid, "palletNumber": pallet_number, "slots": slots}


def _load_styles() -> str:
    if STYLES_PATH.is_file():
        return STYLES_PATH.read_text(encoding="utf-8")
    return "body{margin:0;font-family:Arial,sans-serif;font-size:10pt;} table{border-collapse:collapse;width:100%;}"


def _load_fragment() -> str:
    if FRAGMENT_PATH.is_file():
        return FRAGMENT_PATH.read_text(encoding="utf-8")
    raise RuntimeError(f"Не найден фрагмент шаблона: {FRAGMENT_PATH}")


def _subst(tpl: str, mapping: dict[str, str]) -> str:
    out = tpl
    for k, v in mapping.items():
        out = out.replace("{{" + k + "}}", v)
    return out


def _build_line_rows(items: list[dict[str, Any]], pal: dict[str, Any]) -> tuple[str, float, float, float, float]:
    """HTML строк таблицы, суммы: штуки, коробки, вес (кг), объём (заглушка)."""
    rows: list[str] = []
    tot_pieces = 0.0
    tot_boxes = 0.0
    tot_weight = 0.0
    tot_vol = 0.0
    n = 0
    for slot in pal.get("slots") or []:
        ns = _normalize_slot(slot if isinstance(slot, dict) else {})
        li = ns["lineIndex"]
        if li is None:
            continue
        if li < 0 or li >= len(items):
            continue
        it = items[li]
        alloc = _slot_to_alloc(ns)
        pcs = _allocation_to_pieces(it, alloc)
        box = _allocation_to_boxes(it, alloc)
        if pcs <= 0 and box <= 0:
            continue
        n += 1
        tot_pieces += pcs
        tot_boxes += box
        bw = float(it.get("box_weight") or 0)
        if bw > 0 and box > 0:
            tot_weight += box * bw
        name_e = html.escape((it.get("name") or "—").strip() or "—")
        art_raw = (it.get("article") or "").strip()
        art_e = html.escape(art_raw)
        title = f"{art_e} — {name_e}" if art_raw else name_e
        pib = max(0, int(it.get("pieces_in_box") or 0))
        comment = ""
        if pib > 0 and round(pcs) > 0:
            comment = f"кор({ _fmt_ru_num(pcs, max_decimals=0) })"
        comment_e = html.escape(comment)
        wv = f"{_fmt_ru_num(bw * box, max_decimals=2)}/—" if bw > 0 else "—/—"
        rows.append(
            '<tr class="R12">'
            '<td class="R12C0"><span></span></td>'
            f'<td class="R12C1"><span class="nw">{n}</span></td>'
            f'<td class="R12C2" colspan="3"><span class="sheet-long-text">{title}</span></td>'
            f'<td class="R12C5"><span class="nw">{_fmt_ru_num(pcs, max_decimals=0)}</span></td>'
            f'<td class="R12C5"><span class="nw">{_fmt_ru_num(box, max_decimals=3)}</span></td>'
            f'<td class="R12C7"><span class="nw">{comment_e}</span></td>'
            f'<td class="R12C7"><span class="nw">{html.escape(wv)}</span></td>'
            "<td><span></span></td><td></td>"
            "</tr>"
        )
    if not rows:
        rows.append(
            '<tr class="R12"><td colspan="11" class="R12C0">Нет распределённых позиций на этой паллете.</td></tr>'
        )
    return "".join(rows), tot_pieces, tot_boxes, tot_weight, tot_vol


def build_generic_packing_sheets_html(detail: dict[str, Any]) -> tuple[str | None, str | None]:
    """
    Полный HTML-документ. Ошибка: (None, код_или_текст).
    """
    if _is_arnest_unirus_client(str(detail.get("client") or "")):
        return None, "arnest_client"
    st = detail.get("assemble_state")
    if not isinstance(st, dict):
        return None, "no_assembly"
    pallets_raw = st.get("pallets")
    if not isinstance(pallets_raw, list) or len(pallets_raw) == 0:
        return None, "no_pallets"
    pallets = [_migrate_pallet(p) for p in pallets_raw if isinstance(p, dict)]
    pallets = [p for p in pallets if p.get("slots")]
    if not pallets:
        return None, "no_pallets"
    items = detail.get("items")
    if not isinstance(items, list):
        items = []
    sheet_no = str(detail.get("id") or "")
    ship_ru = _ship_date_ru(str(detail.get("ship_date") or ""))
    buyer = html.escape(str(detail.get("client") or "—").strip() or "—")
    supplier = html.escape(_supplier_default())
    total_pallets = len(pallets)
    try:
        frag = _load_fragment()
    except RuntimeError as e:
        return None, str(e)

    sections: list[str] = []
    for idx, pal in enumerate(pallets, start=1):
        pnum_raw = str(pal.get("palletNumber") or "").strip()
        try:
            pnum_disp = str(int(float(pnum_raw.replace(",", ".")))) if pnum_raw else str(idx)
        except ValueError:
            pnum_disp = pnum_raw or str(idx)
        lines_html, sp, sb, sw, sv = _build_line_rows(items, pal)
        sw_s = _fmt_ru_num(sw, max_decimals=2)
        sv_s = _fmt_ru_num(sv, max_decimals=3) if sv > 0 else "—"
        mapping = {
            "sheet_no": html.escape(sheet_no),
            "ship_date": html.escape(ship_ru),
            "pallet_no": html.escape(pnum_disp),
            "pallets_total": str(total_pallets),
            "supplier": supplier,
            "buyer": buyer,
            "lines_html": lines_html,
            "sum_pieces": _fmt_ru_num(sp, max_decimals=0),
            "sum_boxes": _fmt_ru_num(sb, max_decimals=3),
            "sum_weight": sw_s,
            "sum_volume": sv_s,
        }
        sections.append(
            '<table cellspacing="0" class="generic-pallet-sheet" style="width:100%">'
            + _generic_colgroup_html()
            + _subst(frag, mapping)
            + "</table>"
        )

    styles = _load_styles()
    doc = (
        "<!DOCTYPE html>\n<html lang=\"ru\"><head><meta charset=\"utf-8\"/>"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"/>"
        "<title>Упаковочные листы</title><style>"
        + styles
        + "</style></head><body>\n"
        + "\n".join(sections)
        + "\n</body></html>"
    )
    return doc, None


def build_generic_packing_sheets_pdf_bytes(detail: dict[str, Any]) -> tuple[bytes | None, str | None]:
    """Тот же макет, что HTML, конвертация в PDF (WeasyPrint; на сервере нужны Pango/Cairo)."""
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
    except Exception as exc:  # noqa: BLE001 — отдаём текст на API
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

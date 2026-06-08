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


def _is_lab_industries_client(client: str) -> bool:
    n = str(client or "").strip().replace("  ", " ").lower()
    return n == "лаб индастриз"


def _is_drogeri_retail_client(client: str) -> bool:
    n = str(client or "").strip().replace("  ", " ").lower()
    return n == "дрогери ритейл"


def _ship_date_ru(ship_date: str) -> str:
    s = str(ship_date or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        y, mo, d = m.groups()
        return f"{d}.{mo}.{y}"
    return s or "—"


def _migrate_pallet(p: dict[str, Any]) -> dict[str, Any]:
    """Нормализация id и номера паллеты (как раньше, для совместимости со сборкой)."""
    pid = int(p.get("id") or 0)
    pallet_number = str(p.get("palletNumber") or "").strip()
    if isinstance(p.get("slots"), list):
        return {"id": pid, "palletNumber": pallet_number, "slots": p["slots"]}
    return {"id": pid, "palletNumber": pallet_number, "slots": []}


def _pallet_no_display(pnum_raw: str, pallet_index: int) -> str:
    """Номер паллеты в шапке листа: цифра в рамке; если пусто — порядковый номер."""
    raw = str(pnum_raw or "").strip()
    if not raw:
        return str(pallet_index)
    try:
        return str(int(float(raw.replace(",", "."))))
    except ValueError:
        return raw


def _pallet_sort_key(pal: dict[str, Any], fallback_index: int) -> tuple[int, float, str]:
    """Ключ сортировки: числовой номер паллеты, иначе текст; пустой — порядковый fallback_index."""
    raw = str(pal.get("palletNumber") or pal.get("pallet_number") or "").strip()
    if not raw:
        return (0, float(fallback_index), "")
    try:
        return (0, float(raw.replace(",", ".")), "")
    except ValueError:
        return (1, 0.0, raw.casefold())


def sort_assemble_pallets_by_number(pallets: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Паллеты по возрастанию номера паллеты (для PDF и предпросмотра)."""
    if len(pallets) < 2:
        return list(pallets)
    indexed = list(enumerate(pallets, start=1))
    indexed.sort(key=lambda pair: _pallet_sort_key(pair[1], pair[0]))
    return [p for _, p in indexed]


SUPPLIER_LINE = "М.К. АСЕПТИКА ООО"

# Синхрон с packing_sheet_generic_styles.css: .generic-lines-table font-size
_GENERIC_LINES_TABLE_BASE_PT = 10.89
# Нижняя граница подбора кегля (редкие переполненные листы)
_GENERIC_LINES_TABLE_ABS_MIN_PT = 2.0
# A4 landscape, поля 10 мм: печатная высота ~190 мм
_GENERIC_PAGE_CONTENT_H_MM = 190.0
# Оценка высоты строк 1–4, margin-top у таблицы и запас (pt)
_GENERIC_SHEET_HEADER_RESERVE_PT = 140.0
# Оценка высоты таблицы занижает реальную вёрстку — запас против обрезания
_GENERIC_TABLE_HEIGHT_SAFETY = 1.06


def _mm_to_pt(mm: float) -> float:
    return mm * (72.0 / 25.4)


def _aggregate_pallet_lines(
    items: list[dict[str, Any]], pal: dict[str, Any]
) -> tuple[list[int], dict[int, dict[str, float]]]:
    """Строки таблицы номенклатуры по паллете: порядок lineIndex и агрегаты pcs/boxes."""
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
    return order_keys, agg


def _estimate_generic_table_row_height_pt(font_pt: float, title_len: int) -> float:
    """Грубая оценка высоты строки таблицы (pt) с учётом переноса в колонке «Номенклатура» (~40% ширины)."""
    pad_v_pt = max(0.55, font_pt * 0.32 + 0.45)
    line_h = font_pt * 1.25
    content_w_mm = 297.0 - 20.0
    name_col_pt = _mm_to_pt(content_w_mm) * 0.40
    ch_w = max(font_pt * 0.48, 0.1)
    cpl = max(6, int(name_col_pt / ch_w))
    lines = max(1, (int(title_len) + cpl - 1) // cpl)
    return lines * line_h + 2.0 * pad_v_pt + 2.2


def _estimate_generic_table_block_height_pt(font_pt: float, title_lens: list[int]) -> float:
    """Оценка полной высоты thead + tbody + tfoot (pt)."""
    if not title_lens:
        return _estimate_generic_table_row_height_pt(font_pt, 52)
    h_head = _estimate_generic_table_row_height_pt(font_pt, 22)
    h_foot = _estimate_generic_table_row_height_pt(font_pt, 20)
    h_body = sum(_estimate_generic_table_row_height_pt(font_pt, L) for L in title_lens)
    return h_head + h_body + h_foot


def _fit_generic_table_font_pt(
    title_lens: list[int], *, header_extra_pt: float = 0.0
) -> float:
    """Кегль таблицы: базовый или меньше, чтобы оценочная высота уместилась на одной странице."""
    budget = (
        _mm_to_pt(_GENERIC_PAGE_CONTENT_H_MM)
        - _GENERIC_SHEET_HEADER_RESERVE_PT
        - max(0.0, float(header_extra_pt))
    )
    if budget < 80.0:
        budget = 320.0
    budget_eff = budget / _GENERIC_TABLE_HEIGHT_SAFETY

    def block_h(pt: float) -> float:
        return _estimate_generic_table_block_height_pt(pt, title_lens)

    base = _GENERIC_LINES_TABLE_BASE_PT
    if block_h(base) <= budget_eff:
        return base

    lo, hi = _GENERIC_LINES_TABLE_ABS_MIN_PT, base
    best = _GENERIC_LINES_TABLE_ABS_MIN_PT
    for _ in range(34):
        mid = (lo + hi) / 2.0
        if block_h(mid) <= budget_eff:
            best = mid
            lo = mid
        else:
            hi = mid
    font = max(_GENERIC_LINES_TABLE_ABS_MIN_PT, round(min(base, best), 2))
    while font > _GENERIC_LINES_TABLE_ABS_MIN_PT + 1e-6 and block_h(font) > budget_eff:
        font = max(_GENERIC_LINES_TABLE_ABS_MIN_PT, round(font * 0.985, 2))
    return font


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
    """Колонка «Номенклатура» на листе — только наименование."""
    return (it.get("name") or "—").strip() or "—"


def _drogeri_merge_name_key(name: str) -> str:
    return re.sub(r"\s+", " ", str(name or "").strip()).lower()


def _order_line_max_pieces(it: dict[str, Any]) -> float:
    try:
        q = float(str(it.get("quantity") or "").replace(",", "."))
    except (TypeError, ValueError):
        return 0.0
    if q <= 0:
        return 0.0
    unit = str(it.get("unit") or "piece").strip().lower()
    pib = max(0, int(it.get("pieces_in_box") or 0))
    pps = max(1, int(it.get("pieces_per_set") or 1))
    if unit == "box":
        return float(q * pib) if pib > 0 else q
    if unit == "set":
        return float(q * pps)
    return q


def _pieces_to_order_quantity(it: dict[str, Any], pieces: float) -> float:
    p = max(0.0, float(pieces))
    pib = max(0, int(it.get("pieces_in_box") or 0))
    pps = max(1, int(it.get("pieces_per_set") or 1))
    unit = str(it.get("unit") or "piece").strip().lower()
    if unit == "box":
        return p / pib if pib > 0 else p
    if unit == "set":
        return p / pps
    return p


def _drogeri_merge_source_groups(raw_items: list[dict[str, Any]]) -> list[list[int]]:
    """Индексы исходных строк заказа в каждой объединённой позиции сборки."""
    groups: list[list[int]] = []
    key_to_gi: dict[str, int] = {}
    for idx, raw in enumerate(raw_items or []):
        if not isinstance(raw, dict):
            continue
        name_key = _drogeri_merge_name_key(str(raw.get("name") or ""))
        key = name_key or f"__idx_{idx}"
        gi = key_to_gi.get(key)
        if gi is None:
            gi = len(groups)
            key_to_gi[key] = gi
            groups.append([])
        groups[gi].append(idx)
    return groups


def _drogeri_buyer_order_label(raw_items: list[dict[str, Any]], source_idx: int) -> str:
    if source_idx < 0 or source_idx >= len(raw_items):
        return "—"
    bo = str(raw_items[source_idx].get("buyer_order") or "").strip()
    return bo if bo else "—"


def _drogeri_split_pieces_fifo(
    pieces: float,
    source_indices: list[int],
    raw_items: list[dict[str, Any]],
    consumed: dict[int, float],
) -> list[tuple[str, float]]:
    """Распределяет штуки по номерам заказа в порядке строк заказа (FIFO)."""
    remaining = max(0.0, float(pieces))
    if remaining <= 1e-9 or not source_indices:
        return []
    splits: list[tuple[str, float]] = []
    for si in source_indices:
        if remaining <= 1e-9:
            break
        it = raw_items[si] if 0 <= si < len(raw_items) else {}
        ordered = _order_line_max_pieces(it)
        already = float(consumed.get(si) or 0.0)
        available = max(0.0, ordered - already)
        take = min(remaining, available)
        if take <= 1e-9:
            continue
        bo = _drogeri_buyer_order_label(raw_items, si)
        if splits and splits[-1][0] == bo:
            splits[-1] = (bo, splits[-1][1] + take)
        else:
            splits.append((bo, take))
        consumed[si] = already + take
        remaining -= take
    if remaining > 1e-9:
        si = source_indices[-1]
        bo = _drogeri_buyer_order_label(raw_items, si)
        if splits and splits[-1][0] == bo:
            splits[-1] = (bo, splits[-1][1] + remaining)
        else:
            splits.append((bo, remaining))
        consumed[si] = float(consumed.get(si) or 0.0) + remaining
    return splits


def _drogeri_collect_pallet_splits_by_buyer(
    pal: dict[str, Any],
    merged_items: list[dict[str, Any]],
    source_groups: list[list[int]],
    raw_items: list[dict[str, Any]],
    consumed: dict[int, float],
) -> list[tuple[str, dict[int, dict[str, float]]]]:
    """Группы по номеру заказа на одной физической паллете."""
    by_buyer: dict[str, dict[int, dict[str, float]]] = {}
    buyer_order: list[str] = []
    for slot in pal.get("slots") or []:
        if not isinstance(slot, dict):
            continue
        ns = _normalize_slot(slot)
        li = ns["lineIndex"]
        if li is None or li < 0 or li >= len(merged_items):
            continue
        if li >= len(source_groups):
            continue
        it = merged_items[li]
        pcs = _allocation_to_pieces(it, _slot_to_alloc(ns))
        if pcs <= 1e-9:
            continue
        for bo, take_pcs in _drogeri_split_pieces_fifo(
            pcs, source_groups[li], raw_items, consumed
        ):
            if take_pcs <= 1e-9:
                continue
            pib = max(0, int(it.get("pieces_in_box") or 0))
            take_boxes = take_pcs / pib if pib > 0 else 0.0
            if bo not in by_buyer:
                by_buyer[bo] = {}
                buyer_order.append(bo)
            row = by_buyer[bo].setdefault(li, {"pcs": 0.0, "boxes": 0.0})
            row["pcs"] += take_pcs
            row["boxes"] += take_boxes
    return [(bo, by_buyer[bo]) for bo in buyer_order if by_buyer.get(bo)]


def merge_order_items_for_drogeri(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Сборка/листы: одна строка на наименование, количество суммируется."""
    groups: list[dict[str, Any]] = []
    key_to_gi: dict[str, int] = {}
    for idx, raw in enumerate(items or []):
        if not isinstance(raw, dict):
            continue
        name_key = _drogeri_merge_name_key(str(raw.get("name") or ""))
        key = name_key or f"__idx_{idx}"
        line_buyer = str(raw.get("buyer_order") or "").strip()
        gi = key_to_gi.get(key)
        if gi is None:
            gi = len(groups)
            key_to_gi[key] = gi
            groups.append(
                {
                    "rep": dict(raw),
                    "total_pieces": _order_line_max_pieces(raw),
                    "buyer_orders": [line_buyer] if line_buyer else [],
                }
            )
        else:
            g = groups[gi]
            g["total_pieces"] = float(g["total_pieces"]) + _order_line_max_pieces(raw)
            if line_buyer and line_buyer not in g["buyer_orders"]:
                g["buyer_orders"].append(line_buyer)
    merged: list[dict[str, Any]] = []
    for g in groups:
        rep = dict(g["rep"])
        rep["quantity"] = _pieces_to_order_quantity(rep, float(g["total_pieces"]))
        if g["buyer_orders"]:
            rep["buyer_order"] = ", ".join(g["buyer_orders"])
        merged.append(rep)
    return merged


def _generic_row34_html(label: str, value: str) -> str:
    label_e = html.escape(label, quote=True)
    value_e = html.escape(value, quote=True)
    return (
        '<div class="generic-row34">'
        f'<span class="generic-r34-label">{label_e}</span>'
        f'<span class="generic-r34-center">{value_e}</span>'
        "</div>"
    )


def _generic_row_head_html(ship_e: str, pallet_disp: str, total_pallets: int) -> str:
    pallet_disp_e = html.escape(pallet_disp, quote=True)
    total_e = html.escape(str(max(1, int(total_pallets))), quote=True)
    return (
        '<table class="generic-row-head" role="presentation" aria-label="Номер паллеты и дата">'
        "<tr>"
        '<td class="generic-row-head-left">'
        '<span class="generic-r2-label">Номер паллета</span>'
        "</td>"
        '<td class="generic-row-head-center">'
        '<table class="generic-r2-triplet" role="presentation" aria-label="Номер паллеты из общего числа">'
        "<tr>"
        f'<td class="generic-r2-cell generic-r2-cell--side">{pallet_disp_e}</td>'
        '<td class="generic-r2-cell generic-r2-cell--mid">из</td>'
        f'<td class="generic-r2-cell generic-r2-cell--side">{total_e}</td>'
        "</tr></table>"
        "</td>"
        '<td class="generic-row-head-right">'
        '<span class="generic-r1-label">Дата</span> '
        f'<span class="generic-r1-frame">{ship_e}</span>'
        "</td>"
        "</tr></table>"
    )


def _build_table_html_from_line_agg(
    items: list[dict[str, Any]],
    order_keys: list[int],
    agg: dict[int, dict[str, float]],
    *,
    header_extra_pt: float = 0.0,
) -> str:
    """Таблица номенклатуры по готовому агрегату pcs/boxes."""
    title_lens = [len(_nomenclature_title(items[li])) for li in order_keys]
    table_font_pt = _fit_generic_table_font_pt(
        title_lens, header_extra_pt=header_extra_pt
    )
    fs_attr = ""
    if table_font_pt + 0.005 < _GENERIC_LINES_TABLE_BASE_PT:
        fs_attr = f' style="font-size:{table_font_pt:.2f}pt;line-height:1.25"'

    head = (
        f"<table class=\"generic-lines-table\" cellspacing=\"0\"{fs_attr}>"
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


def _build_pallet_lines_table_html(
    items: list[dict[str, Any]],
    pal: dict[str, Any],
    *,
    header_extra_pt: float = 0.0,
) -> str:
    """Таблица строки 5: агрегат по строкам заказа (lineIndex) на паллете."""
    order_keys, agg = _aggregate_pallet_lines(items, pal)
    return _build_table_html_from_line_agg(
        items, order_keys, agg, header_extra_pt=header_extra_pt
    )


def _build_drogeri_packing_sheet_sections(
    *,
    pallets: list[dict[str, Any]],
    raw_items: list[dict[str, Any]],
    merged_items: list[dict[str, Any]],
    client_name: str,
    ship_e: str,
) -> list[str]:
    """Один лист на пару (физическая паллета, номер заказа); «N из 1» в шапке."""
    source_groups = _drogeri_merge_source_groups(raw_items)
    consumed: dict[int, float] = {}
    drogeri_header_extra_pt = 28.0
    sections: list[str] = []
    for idx, pal in enumerate(pallets, start=1):
        pnum_raw = str(pal.get("palletNumber") or "").strip()
        pallet_disp = _pallet_no_display(pnum_raw, idx)
        splits = _drogeri_collect_pallet_splits_by_buyer(
            pal, merged_items, source_groups, raw_items, consumed
        )
        for buyer_no, line_agg in splits:
            order_keys = [
                li
                for li in line_agg
                if float(line_agg[li].get("pcs") or 0.0) > 1e-9
                or float(line_agg[li].get("boxes") or 0.0) > 1e-9
            ]
            if not order_keys:
                continue
            row_head = _generic_row_head_html(ship_e, pallet_disp, 1)
            row3 = _generic_row34_html("Поставщик:", SUPPLIER_LINE)
            row4 = _generic_row34_html("Покупатель:", client_name)
            row4_order = _generic_row34_html("Номер заказа:", buyer_no)
            row5 = (
                '<div class="generic-row5-wrap">'
                + _build_table_html_from_line_agg(
                    merged_items,
                    order_keys,
                    line_agg,
                    header_extra_pt=drogeri_header_extra_pt,
                )
                + "</div>"
            )
            sections.append(
                f'<div class="generic-pallet-sheet">{row_head}{row3}{row4}{row4_order}{row5}</div>'
            )
    return sections


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
    if _is_lab_industries_client(str(detail.get("client") or "")):
        return None, "lab_client"
    st = detail.get("assemble_state")
    if not isinstance(st, dict):
        return None, "no_assembly"
    pallets_raw = st.get("pallets")
    if not isinstance(pallets_raw, list) or len(pallets_raw) == 0:
        return None, "no_pallets"
    pallets = sort_assemble_pallets_by_number(
        [_migrate_pallet(p) for p in pallets_raw if isinstance(p, dict)]
    )
    if not pallets:
        return None, "no_pallets"

    raw_items = detail.get("items")
    if not isinstance(raw_items, list):
        raw_items = []
    client_name = str(detail.get("client") or "").strip() or "—"
    is_drogeri = _is_drogeri_retail_client(client_name)

    ship_ru = _ship_date_ru(str(detail.get("ship_date") or ""))
    ship_e = html.escape(ship_ru, quote=True)

    if is_drogeri:
        merged_items = merge_order_items_for_drogeri(raw_items)
        sections = _build_drogeri_packing_sheet_sections(
            pallets=pallets,
            raw_items=raw_items,
            merged_items=merged_items,
            client_name=client_name,
            ship_e=ship_e,
        )
        if not sections:
            return None, "no_pallets"
    else:
        items = raw_items
        total_pallets = len(pallets)
        sections = []
        for idx, pal in enumerate(pallets, start=1):
            pnum_raw = str(pal.get("palletNumber") or "").strip()
            pallet_disp = _pallet_no_display(pnum_raw, idx)
            row_head = _generic_row_head_html(ship_e, pallet_disp, total_pallets)
            row3 = _generic_row34_html("Поставщик:", SUPPLIER_LINE)
            row4 = _generic_row34_html("Покупатель:", client_name)
            row5 = (
                '<div class="generic-row5-wrap">'
                + _build_pallet_lines_table_html(items, pal)
                + "</div>"
            )
            sections.append(
                f'<div class="generic-pallet-sheet">{row_head}{row3}{row4}{row5}</div>'
            )

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
        "lab_client": "Для клиента «ЛАБ Индастриз» используйте отдельные паллетные листы (GS1-128).",
        "no_assembly": "Нет данных сборки (assemble_state). Сохраните сборку в разделе «Сборка заказа».",
        "no_pallets": "Нет паллет в сборке.",
        "no_pdf_engine": (
            "Не удалось загрузить WeasyPrint (HTML→PDF). На Linux установите системные "
            "библиотеки Pango/Cairo (см. документацию WeasyPrint) и пакет weasyprint из requirements.txt."
        ),
        "pdf_empty": "Сформирован пустой PDF.",
    }.get(code, str(code))

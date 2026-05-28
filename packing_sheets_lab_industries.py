# -*- coding: utf-8 -*-
"""Паллетные листы для клиента «ЛАБ Индастриз» (PDF, GS1-128 SSCC для номера паллеты)."""

from __future__ import annotations

import io
import re
from typing import Any

_LAB_CLIENT_NORMALIZED = "лаб индастриз"

# A4 портрет.
_LAB_A4_W_MM = 210.0
_LAB_MARGIN_MM = 12.0
_LAB_BOX_PAD_MM = 3.0
_LAB_ROW1_TOP_MM = 14.0
_LAB_TEXT_FONT_PT = 12.0
_LAB_TEXT_H_MM = 7.0
_LAB_TAB_MM = 12.5
_LAB_BARCODE_W_MM = 72.0
_LAB_BARCODE_HEIGHT_SCALE = 0.55
_LAB_ROW1_PALLET_BARCODE_HEIGHT_SCALE = _LAB_BARCODE_HEIGHT_SCALE * 1.8
_LAB_BARCODE_CAPTION_GAP_MM = 1.5
_LAB_BARCODE_CAPTION_LINE_H_MM = 6.0
_LAB_MIN_RIGHT_W_MM = 95.0
MAX_PALLET_SHEET_PDF_PAGES = 500


def is_lab_industries_client(client: str) -> bool:
    n = str(client or "").strip().replace("  ", " ").lower()
    return n == _LAB_CLIENT_NORMALIZED


def _parse_pallet_number_int(pallet_number: str, fallback_index: int) -> int:
    raw = str(pallet_number).strip()
    if not raw:
        return max(1, int(fallback_index))
    try:
        return max(1, int(float(raw.replace(",", "."))))
    except ValueError:
        digits = re.sub(r"\D", "", raw)
        if digits:
            try:
                return max(1, int(digits))
            except ValueError:
                pass
    return max(1, int(fallback_index))


def build_lab_pallet_sscc_gs1(
    pallet_number: str,
    fallback_index: int = 1,
    *,
    ai_seq: int = 0,
) -> tuple[str, str]:
    """GS1-128: префикс AI по заказу (01)…(99), затем (00); SSCC 18 цифр + суффикс паллеты."""
    n = _parse_pallet_number_int(pallet_number, fallback_index)
    sscc18 = f"150000000000000{n:03d}"
    if len(sscc18) != 18:
        sscc18 = (sscc18 + "0" * 18)[:18]
    ai = int(ai_seq) % 100
    ai_str = f"{ai:02d}"
    encode = ai_str + sscc18
    human = f"({ai_str}){sscc18}"
    return encode, human


def _first_line_on_pallet(
    pal: dict[str, Any], items: list[dict[str, Any]]
) -> tuple[int | None, str, str]:
    """lineIndex, артикул и наименование первой позиции на паллете."""
    for slot in pal.get("slots") or []:
        if not isinstance(slot, dict):
            continue
        li = slot.get("lineIndex")
        if li is None or li == "":
            continue
        try:
            idx = int(li)
        except (TypeError, ValueError):
            continue
        if 0 <= idx < len(items):
            it = items[idx]
            art = str(it.get("article") or "").strip()
            name = str(it.get("name") or "").strip()
            if art or name:
                return idx, art or "—", name or "—"
    return None, "—", "—"


def _pieces_by_line_index(
    items: list[dict[str, Any]], pallets_raw: list[dict[str, Any]]
) -> dict[int, float]:
    """Сумма штук по lineIndex по всем паллетам сборки."""
    from packing_sheets_generic import (
        _allocation_to_pieces,
        _normalize_slot,
        _slot_to_alloc,
    )

    totals: dict[int, float] = {}
    for pal in pallets_raw:
        if not isinstance(pal, dict):
            continue
        for slot in pal.get("slots") or []:
            if not isinstance(slot, dict):
                continue
            ns = _normalize_slot(slot)
            li = ns["lineIndex"]
            if li is None or li < 0 or li >= len(items):
                continue
            pcs = _allocation_to_pieces(items[li], _slot_to_alloc(ns))
            totals[li] = totals.get(li, 0.0) + pcs
    return totals


_LAB_ROW4_LABEL = "Количество упаковок на поддоне, шт"
_LAB_ROW4_KOR_SUFFIX = " кор."
_LAB_ROW4_MIN_RIGHT_W_MM = 36.0
_LAB_ROW5_LABEL = "Номер заказа:"
_LAB_ROW6_LABEL = "Общее количество заказа:"
_LAB_ROW7_LABEL = "Номер паллеты:"
_LAB_ROW7_VALUE_FONT_PT = _LAB_TEXT_FONT_PT * 2.0
_LAB_ROW7_VALUE_H_MM = _LAB_TEXT_H_MM * 2.0
_LAB_ROW8_MANUFACTURER_TEXT = (
    "Изготовитель: (K) Общество с ограниченной ответственностью «ЛАБ Индастриз», "
    "107045, Россия, г. Москва, Колокольников пер., д.11. "
    "(T) Общество с ограниченной ответственностью «ЛАБ Индастриз», 123112, Россия, "
    "г. Москва, вн.тер.г. муниципальный округ Пресненский, ул. Тестовская, д.10, "
    "помещ.1/16."
)
_LAB_ROW9_PRODUCTION_ADDRESS_TEXT = (
    "Адрес производства: 140032, Россия, Московская область, г.о. Люберцы, "
    "р.п. Малаховка, ул. Шоссейная, д. 40, ООО «М.К. Асептика»"
)
_LAB_ROW10_BATCH_LABEL = "Номер партии:"
_LAB_ROW10_EXPIRY_LABEL = "Годен до:"


def _lab_assemble_fields_from_pallet(pal: dict[str, Any]) -> tuple[str, str]:
    """Партия и «Годен до» из первого слота паллеты с заполненными полями."""
    batch_raw = ""
    expiry_raw = ""
    for slot in pal.get("slots") or []:
        if not isinstance(slot, dict):
            continue
        bn = str(slot.get("batchNumber") or "").strip()
        le = str(slot.get("labExpiryDate") or "").strip()
        if bn and not batch_raw:
            batch_raw = bn
        if le and not expiry_raw:
            expiry_raw = le
    return batch_raw, expiry_raw


def _lab_normalize_batch_single(raw: str) -> str:
    """Одна партия на строку товара на паллете (без перечисления через запятую)."""
    s = str(raw or "").strip()
    if not s:
        return ""
    return s.split(",")[0].strip()


def _lab_batch_parts(raw: str, max_parts: int = 1) -> list[str]:
    part = _lab_normalize_batch_single(raw)
    return [part] if part else []


def _format_lab_single_batch_display(part: str) -> str:
    """ТA127361601 → ТA 127361601."""
    part = str(part or "").strip()
    if not part:
        return "—"
    if re.search(r"\s", part):
        return part
    m = re.match(r"^([^\d\s/]+)(\d[\d\s]*)$", part, re.UNICODE)
    if m:
        return f"{m.group(1).rstrip()} {m.group(2).lstrip()}"
    return part


def _format_lab_row10_batch_lines(raw: str) -> list[str]:
    """Одна партия на паллетный лист."""
    part = _lab_normalize_batch_single(raw)
    if not part:
        return ["—"]
    return [_format_lab_single_batch_display(part)]


def _format_lab_row10_batch(raw: str) -> str:
    return "\n".join(_format_lab_row10_batch_lines(raw))


def _lab_draw_row10_batch_block(
    pdf: Any,
    x: float,
    y: float,
    inner_w: float,
    batch_lines: list[str],
    h_txt: float,
    fs: float,
    tab_mm: float,
) -> None:
    pdf.set_font("PLCalibri", "B", fs)
    w_label = pdf.get_string_width(_LAB_ROW10_BATCH_LABEL)
    value_x = x + w_label + tab_mm
    value_w = max(1.0, inner_w - w_label - tab_mm)
    block_h = max(h_txt, len(batch_lines) * h_txt)
    label_y = y + (block_h - h_txt) / 2.0
    pdf.set_xy(x, y)
    pdf.cell(w_label, h_txt, _LAB_ROW10_BATCH_LABEL, align="L")
    for i, line in enumerate(batch_lines):
        pdf.set_xy(value_x, y + i * h_txt)
        pdf.cell(value_w, h_txt, line, align="L")


def _format_lab_row10_expiry(raw: str) -> str:
    """ДДММГГ (030428) → 04/2028; также 0428, 04/2028."""
    s = str(raw or "").strip()
    if not s:
        return "—"
    digits = re.sub(r"\D", "", s)
    if len(digits) >= 6:
        d = digits[-6:]
        return f"{d[2:4]}/20{d[4:6]}"
    if "/" in s:
        parts = [p.strip() for p in s.split("/", 1)]
        if len(parts) == 2:
            mm_d = re.sub(r"\D", "", parts[0])
            yy_d = re.sub(r"\D", "", parts[1])
            if mm_d and yy_d:
                mm = mm_d[:2].zfill(2)
                yy = yy_d if len(yy_d) == 4 else ("20" + yy_d.zfill(2)) if len(yy_d) <= 2 else yy_d
                if len(yy) == 4:
                    return f"{mm}/{yy}"
    if len(digits) == 4:
        return f"{digits[:2]}/20{digits[2:]}"
    return s


def _lab_batch_has_cyrillic(raw: str) -> bool:
    return bool(re.search(r"[а-яёА-ЯЁ]", str(raw or "")))


def _format_lab_batch_for_barcode(raw: str) -> str:
    """ТA127361601 → A127361601 (последняя буква префикса + цифры); первая партия."""
    parts = _lab_batch_parts(raw, 1)
    part = parts[0] if parts else ""
    if not part:
        return ""
    compact = re.sub(r"\s+", "", part)
    m = re.match(r"^([^\d]+)(\d+)$", compact, re.UNICODE)
    if m:
        letters, digits = m.group(1), m.group(2)
        if letters:
            return letters[-1] + digits
        return digits
    return compact


def build_lab_product_gs1_128(
    article: str, batch_raw: str, expiry_raw: str
) -> tuple[str | None, str | None, str | None]:
    """GS1-128: (240)артикул(10)партия(15)YYMMDD; ошибка в третьем элементе."""
    from api_server import _arnest_first_digit_run, _arnest_yymmdd_for_barcode

    art_digits = _arnest_first_digit_run(str(article or "").strip())
    if not art_digits:
        return None, None, "в артикуле нет цифр"
    if _lab_batch_has_cyrillic(str(batch_raw or "")):
        return None, None, "в номере партии нельзя использовать русские буквы"
    batch_bc = _format_lab_batch_for_barcode(batch_raw)
    if not batch_bc:
        return None, None, "не указана партия"
    exp_bc = _arnest_yymmdd_for_barcode(str(expiry_raw or "").strip())
    if not exp_bc:
        return None, None, "срок годности: укажите 6 цифр ДДММГГ (например 030428)"
    gs1 = f"240{art_digits}\x1d10{batch_bc}\x1d15{exp_bc}"
    hri = f"(240){art_digits}(10){batch_bc}(15){exp_bc}"
    return gs1, hri, None


def _lab_draw_bold_label_tab_value(
    pdf: Any,
    x: float,
    y: float,
    inner_w: float,
    label: str,
    value: str,
    h_txt: float,
    fs: float,
    tab_mm: float,
) -> None:
    pdf.set_font("PLCalibri", "B", fs)
    w_label = pdf.get_string_width(label)
    pdf.set_xy(x, y)
    pdf.cell(w_label, h_txt, label, align="L")
    rest_w = max(1.0, inner_w - w_label - tab_mm)
    pdf.set_xy(x + w_label + tab_mm, y)
    pdf.cell(rest_w, h_txt, value or "—", align="L")


def _boxes_on_pallet(
    pal: dict[str, Any], items: list[dict[str, Any]]
) -> float:
    """Сумма коробок по слотам на одной паллете."""
    from packing_sheets_generic import (
        _allocation_to_boxes,
        _normalize_slot,
        _slot_to_alloc,
    )

    total = 0.0
    for slot in pal.get("slots") or []:
        if not isinstance(slot, dict):
            continue
        ns = _normalize_slot(slot)
        li = ns["lineIndex"]
        if li is None or li < 0 or li >= len(items):
            continue
        total += _allocation_to_boxes(items[li], _slot_to_alloc(ns))
    return total


def _pieces_on_slot(items: list[dict[str, Any]], slot: dict[str, Any]) -> float:
    from packing_sheets_generic import (
        _allocation_to_pieces,
        _normalize_slot,
        _slot_to_alloc,
    )

    ns = _normalize_slot(slot)
    li = ns["lineIndex"]
    if li is None or li < 0 or li >= len(items):
        return 0.0
    return float(_allocation_to_pieces(items[li], _slot_to_alloc(ns)) or 0)


def _boxes_on_slot(items: list[dict[str, Any]], slot: dict[str, Any]) -> float:
    from packing_sheets_generic import (
        _allocation_to_boxes,
        _normalize_slot,
        _slot_to_alloc,
    )

    ns = _normalize_slot(slot)
    li = ns["lineIndex"]
    if li is None or li < 0 or li >= len(items):
        return 0.0
    return float(_allocation_to_boxes(items[li], _slot_to_alloc(ns)) or 0)


def _lab_productive_slots(
    pal: dict[str, Any], items: list[dict[str, Any]]
) -> list[tuple[dict[str, Any], int]]:
    """Слоты паллеты с выбранной позицией заказа — по одному листу PDF на слот."""
    from packing_sheets_generic import _normalize_slot

    out: list[tuple[dict[str, Any], int]] = []
    for slot in pal.get("slots") or []:
        if not isinstance(slot, dict):
            continue
        ns = _normalize_slot(slot)
        li = ns["lineIndex"]
        if li is None or li < 0 or li >= len(items):
            continue
        out.append((slot, li))
    return out


def _format_lab_row7_pallet_seq(pallet_index: int, pallet_total: int) -> str:
    """Например: 1 из 10."""
    total = max(1, int(pallet_total))
    num = max(1, min(int(pallet_index), total))
    return f"{num} из {total}"


def _format_lab_row6_total_order(qty: float | None) -> str:
    """Например: 20000 шт."""
    if qty is None:
        return "—"
    try:
        n = float(qty)
    except (TypeError, ValueError):
        return "—"
    if n < 0:
        return "—"
    if abs(n - round(n)) < 1e-6:
        n_s = str(int(round(n)))
    else:
        from packing_sheets_generic import _fmt_ru_num  # noqa: PLC0415

        n_s = _fmt_ru_num(n, max_decimals=3)
    return f"{n_s} шт"


def _format_lab_row4_boxes(boxes: float) -> str:
    """Например: 399 кор."""
    if boxes <= 0:
        n_s = "0"
    elif abs(boxes - round(boxes)) < 1e-6:
        n_s = str(int(round(boxes)))
    else:
        from packing_sheets_generic import _fmt_ru_num  # noqa: PLC0415

        n_s = _fmt_ru_num(boxes, max_decimals=3)
    return f"{n_s}{_LAB_ROW4_KOR_SUFFIX}"


def _buyer_order_for_pallet(
    line_idx: int | None,
    items: list[dict[str, Any]],
    order_buyer_order: str,
    buyer_order_mode: str,
) -> str:
    """Номер заказа покупателя по строке заказа; иначе общий (старые заказы)."""
    _ = buyer_order_mode
    if line_idx is not None and 0 <= line_idx < len(items):
        line_bo = str(items[line_idx].get("buyer_order") or "").strip()
        if line_bo:
            return line_bo
    order_bo = str(order_buyer_order or "").strip()
    return order_bo if order_bo else "—"


def _format_lab_row3_text(
    total_pieces: float, volume_ml: float, volume_unit: str = "ml"
) -> str:
    """Например: 2394 х 200мл или 2394 х 70г."""
    pcs = max(0, int(round(total_pieces)))
    vol = float(volume_ml or 0)
    if vol <= 0:
        return str(pcs)
    if abs(vol - round(vol)) < 1e-6:
        vol_s = str(int(round(vol)))
    else:
        vol_s = ("%g" % vol).replace(".", ",")
    unit = (volume_unit or "ml").strip().lower()
    suffix = "г" if unit in ("g", "gr", "г", "гр") else "мл"
    return f"{pcs} х {vol_s}{suffix}"


def lab_pallets_from_order_detail(detail: dict[str, Any]) -> list[dict[str, Any]]:
    """Строки паллет из assemble_state заказа."""
    from packing_sheets_generic import sort_assemble_pallets_by_number

    st = detail.get("assemble_state")
    if not isinstance(st, dict):
        return []
    pallets_raw = st.get("pallets")
    if not isinstance(pallets_raw, list):
        return []
    items = detail.get("items")
    if not isinstance(items, list):
        items = []
    pallets = sort_assemble_pallets_by_number(
        [p for p in pallets_raw if isinstance(p, dict)]
    )
    totals = _pieces_by_line_index(items, pallets)
    order_bo = str(detail.get("buyer_order") or "").strip()
    buyer_mode = str(detail.get("buyer_order_mode") or "single").strip()
    lab_sscc_ai_seq = int(detail.get("lab_sscc_ai_seq") or 0) % 100
    out: list[dict[str, Any]] = []
    pallet_total = len(pallets)
    for idx, pal in enumerate(pallets, start=1):
        pnum = str(pal.get("palletNumber") or "").strip()
        if not pnum:
            pnum = str(idx)
        row7_pallet = _format_lab_row7_pallet_seq(idx, pallet_total)
        productive = _lab_productive_slots(pal, items)
        if not productive:
            line_idx, article, name = _first_line_on_pallet(pal, items)
            total_pcs = totals.get(line_idx, 0.0) if line_idx is not None else 0.0
            boxes_on_pal = _boxes_on_pallet(pal, items)
            vol = 0.0
            vol_unit = "ml"
            if line_idx is not None and 0 <= line_idx < len(items):
                vol = float(items[line_idx].get("volume_ml") or 0)
                vol_unit = str(items[line_idx].get("volume_unit") or "ml")
            buyer_order = _buyer_order_for_pallet(line_idx, items, order_bo, buyer_mode)
            total_order_qty = None
            if line_idx is not None and 0 <= line_idx < len(items):
                raw_tq = items[line_idx].get("total_order_quantity")
                if raw_tq is not None:
                    try:
                        total_order_qty = float(raw_tq)
                    except (TypeError, ValueError):
                        total_order_qty = None
            batch_raw, expiry_raw = _lab_assemble_fields_from_pallet(pal)
            _gs1_data, gs1_hri, _gs1_err = build_lab_product_gs1_128(
                article, batch_raw, expiry_raw
            )
            out.append(
                {
                    "article": article,
                    "name": name,
                    "pallet_number": pnum,
                    "total_pieces": total_pcs,
                    "volume_ml": vol,
                    "volume_unit": vol_unit,
                    "row3_text": _format_lab_row3_text(total_pcs, vol, vol_unit),
                    "boxes_on_pallet": boxes_on_pal,
                    "row4_text": _format_lab_row4_boxes(boxes_on_pal),
                    "buyer_order": buyer_order,
                    "total_order_quantity": total_order_qty,
                    "row6_text": _format_lab_row6_total_order(total_order_qty),
                    "row7_text": row7_pallet,
                    "batch_number": batch_raw,
                    "lab_expiry_date": expiry_raw,
                    "row10_batch_text": _format_lab_row10_batch(batch_raw),
                    "row10_expiry_text": _format_lab_row10_expiry(expiry_raw),
                    "row11_gs1_hri": gs1_hri or "",
                    "lab_sscc_ai_seq": lab_sscc_ai_seq,
                }
            )
            continue
        for _sheet_i, (slot, line_idx) in enumerate(productive, start=1):
            it = items[line_idx]
            article = str(it.get("article") or "").strip() or "—"
            name = str(it.get("name") or "").strip() or "—"
            slot_pcs = _pieces_on_slot(items, slot)
            boxes_on_pal = _boxes_on_slot(items, slot)
            vol = float(it.get("volume_ml") or 0)
            vol_unit = str(it.get("volume_unit") or "ml")
            buyer_order = _buyer_order_for_pallet(line_idx, items, order_bo, buyer_mode)
            total_order_qty = None
            raw_tq = it.get("total_order_quantity")
            if raw_tq is not None:
                try:
                    total_order_qty = float(raw_tq)
                except (TypeError, ValueError):
                    total_order_qty = None
            batch_raw = _lab_normalize_batch_single(str(slot.get("batchNumber") or ""))
            expiry_raw = str(slot.get("labExpiryDate") or "").strip()
            _gs1_data, gs1_hri, _gs1_err = build_lab_product_gs1_128(
                article, batch_raw, expiry_raw
            )
            out.append(
                {
                    "article": article,
                    "name": name,
                    "pallet_number": pnum,
                    "total_pieces": slot_pcs,
                    "volume_ml": vol,
                    "volume_unit": vol_unit,
                    "row3_text": _format_lab_row3_text(slot_pcs, vol, vol_unit),
                    "boxes_on_pallet": boxes_on_pal,
                    "row4_text": _format_lab_row4_boxes(boxes_on_pal),
                    "buyer_order": buyer_order,
                    "total_order_quantity": total_order_qty,
                    "row6_text": _format_lab_row6_total_order(total_order_qty),
                    "row7_text": row7_pallet,
                    "batch_number": batch_raw,
                    "lab_expiry_date": expiry_raw,
                    "row10_batch_text": _format_lab_row10_batch(batch_raw),
                    "row10_expiry_text": _format_lab_row10_expiry(expiry_raw),
                    "row11_gs1_hri": gs1_hri or "",
                    "lab_sscc_ai_seq": lab_sscc_ai_seq,
                }
            )
    return out


def lab_pallet_pdf_error_message(code: str) -> str:
    return {
        "no_fpdf": "На сервере не установлен пакет fpdf2 (pip install fpdf2).",
        "no_barcode": (
            "На сервере не установлен python-barcode[images] "
            '(pip install "python-barcode[images]").'
        ),
        "no_font": "Не найден шрифт Calibri для паллетных листов (каталог fonts/).",
        "validation_pallets": "Укажите от 1 до 500 паллет.",
        "no_pallets": "Нет паллет в сборке.",
        "no_assembly": "Нет данных сборки (assemble_state).",
        "lab_client": "Заказ не для клиента «ЛАБ Индастриз».",
        "pdf_build": "Не удалось собрать PDF.",
        "barcode_fetch": "Не удалось сформировать штрих-код GS1-128.",
    }.get(code, str(code))


def _lab_measure_wrapped_height_mm(
    pdf: Any, text: str, width_mm: float, line_h_mm: float, *, bold: bool
) -> float:
    """Высота блока текста при переносе по ширине (мм)."""
    style = "B" if bold else ""
    pdf.set_font("PLCalibri", style, _LAB_TEXT_FONT_PT)
    words = (text or "—").split()
    if not words:
        return line_h_mm
    lines = 1
    line = ""
    for word in words:
        trial = (line + " " + word).strip()
        if pdf.get_string_width(trial) <= width_mm:
            line = trial
        else:
            if line:
                lines += 1
            line = word
            if pdf.get_string_width(line) > width_mm:
                lines += 1
                line = ""
    return max(1, lines) * line_h_mm


def _lab_row4_frame_widths_mm(
    pdf: Any, content_w: float, pad: float, fs: float, row4_text: str
) -> tuple[float, float]:
    """Ширина левой рамки по подписи; правая — остаток, не уже значения «N кор.»."""
    pdf.set_font("PLCalibri", "B", fs)
    w_label = pdf.get_string_width(_LAB_ROW4_LABEL)
    w_val = pdf.get_string_width(str(row4_text or "0 кор.").strip() or "0 кор.")
    left_w = pad + w_label + pad
    min_right = max(_LAB_ROW4_MIN_RIGHT_W_MM, pad + w_val + pad)
    if left_w > content_w - min_right:
        left_w = content_w - min_right
    left_w = max(left_w, pad + w_label + pad)
    right_w = content_w - left_w
    return left_w, right_w


def _lab_row1_left_width_mm(
    pdf: Any, article: str, content_w: float, pad: float, fs: float, h_txt: float
) -> tuple[float, str, float, float]:
    """Ширина левой рамки по содержимому; артикул для отрисовки (возможно усечён)."""
    label_art = "АРТИКУЛ:"
    pdf.set_font("PLCalibri", "B", fs)
    w_label = pdf.get_string_width(label_art)
    max_art_w = max(
        8.0,
        content_w - _LAB_MIN_RIGHT_W_MM - 2 * pad - w_label - _LAB_TAB_MM - 2 * pad,
    )
    art_show = article
    if pdf.get_string_width(article) > max_art_w:
        from api_server import _arnest_clip_text_to_width_mm  # noqa: PLC0415

        art_show = _arnest_clip_text_to_width_mm(pdf, article, max_art_w) or "—"
    w_article = pdf.get_string_width(art_show)
    left_w = pad + w_label + _LAB_TAB_MM + w_article + pad
    left_w = min(left_w, content_w - _LAB_MIN_RIGHT_W_MM)
    left_w = max(left_w, pad + w_label + pad + 12.0)
    return left_w, art_show, w_label, w_article


def build_lab_industries_pallet_sheets_pdf_bytes(
    pallets: list[dict[str, Any]],
) -> tuple[bytes | None, str | None, str | None]:
    """PDF: строки 1–11 (артикул … партия / годен до / GS1-128 товара)."""
    from api_server import (  # noqa: PLC0415
        FPDF,
        HAVE_FPDF,
        WrapMode,
        _arnest_barcode_caption_font_pt,
        _arnest_clip_text_to_width_mm,
        _arnest_pdf_register_text_fonts,
        _barcode_raster_pixel_size,
        render_gs1_128_barcode_png,
    )

    if not HAVE_FPDF or FPDF is None:
        return None, "no_fpdf", lab_pallet_pdf_error_message("no_fpdf")
    rows = [p for p in pallets if isinstance(p, dict)]
    n = len(rows)
    if n < 1 or n > MAX_PALLET_SHEET_PDF_PAGES:
        return None, "validation_pallets", lab_pallet_pdf_error_message("validation_pallets")
    try:
        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.set_auto_page_break(False)
        font_err = _arnest_pdf_register_text_fonts(pdf)
        if font_err:
            return None, "no_font", lab_pallet_pdf_error_message("no_font")
        content_w = _LAB_A4_W_MM - 2 * _LAB_MARGIN_MM
        fs = _LAB_TEXT_FONT_PT
        h_txt = _LAB_TEXT_H_MM
        pad = _LAB_BOX_PAD_MM
        x0 = _LAB_MARGIN_MM
        wm = WrapMode.WORD if WrapMode is not None else None
        label_pal = "Номер паллета:"

        for idx, row in enumerate(rows, start=1):
            article = str(row.get("article") or "").strip() or "—"
            product_name = str(row.get("name") or "").strip() or "—"
            row3_text = str(row.get("row3_text") or "").strip()
            if not row3_text:
                total_pcs = float(row.get("total_pieces") or 0)
                vol = float(row.get("volume_ml") or 0)
                vol_unit = str(row.get("volume_unit") or "ml")
                row3_text = _format_lab_row3_text(total_pcs, vol, vol_unit)
            row4_text = str(row.get("row4_text") or "").strip()
            if not row4_text:
                boxes_on_pal = float(row.get("boxes_on_pallet") or 0)
                row4_text = _format_lab_row4_boxes(boxes_on_pal)
            pallet_number = str(row.get("pallet_number") or "").strip()
            ai_seq = int(row.get("lab_sscc_ai_seq") or 0) % 100
            gs1_data, gs1_hri = build_lab_pallet_sscc_gs1(
                pallet_number, idx, ai_seq=ai_seq
            )
            try:
                png = render_gs1_128_barcode_png(gs1_data, write_text=False)
            except Exception as exc:
                return (
                    None,
                    "barcode_fetch",
                    f"Паллета {idx}: {exc}",
                )
            dims = _barcode_raster_pixel_size(png)
            if not dims:
                return None, "pdf_build", lab_pallet_pdf_error_message("pdf_build")
            pw, ph = dims

            left_w, art_show, w_label, w_article = _lab_row1_left_width_mm(
                pdf, article, content_w, pad, fs, h_txt
            )
            right_w = content_w - left_w
            x_right = x0 + left_w

            pdf.set_font("PLCalibri", "", fs)
            w_pl = pdf.get_string_width(label_pal)
            right_inner = right_w - 2 * pad
            bc_w = min(
                _LAB_BARCODE_W_MM,
                max(36.0, right_inner - w_pl - _LAB_TAB_MM),
            )
            bc_h = bc_w * (ph / pw) * _LAB_ROW1_PALLET_BARCODE_HEIGHT_SCALE
            cap_fs = _arnest_barcode_caption_font_pt(pdf, gs1_hri, bc_w)
            cap_h = _LAB_BARCODE_CAPTION_GAP_MM + _LAB_BARCODE_CAPTION_LINE_H_MM
            bc_row_h = max(h_txt, bc_h)
            right_inner_h = pad + bc_row_h + cap_h + pad
            left_inner_h = pad + h_txt + pad
            row1_h = max(left_inner_h, right_inner_h)

            row2_inner_w = content_w - 2 * pad
            row2_text_h = _lab_measure_wrapped_height_mm(
                pdf, product_name, row2_inner_w, h_txt, bold=True
            )
            row2_h = pad + row2_text_h + pad
            row3_h = pad + h_txt + pad
            row4_h = pad + h_txt + pad
            row4_left_w, row4_right_w = _lab_row4_frame_widths_mm(
                pdf, content_w, pad, fs, row4_text
            )
            x_row4_right = x0 + row4_left_w
            buyer_order = str(row.get("buyer_order") or "").strip() or "—"
            row5_h = pad + h_txt + pad
            row5_left_w = row4_left_w
            row5_right_w = row4_right_w
            x_row5_right = x_row4_right
            row6_text = str(row.get("row6_text") or "").strip()
            if not row6_text:
                raw_tq = row.get("total_order_quantity")
                tq_val: float | None
                try:
                    tq_val = float(raw_tq) if raw_tq is not None else None
                except (TypeError, ValueError):
                    tq_val = None
                row6_text = _format_lab_row6_total_order(tq_val)
            row6_h = pad + h_txt + pad
            row6_left_w = row4_left_w
            row6_right_w = row4_right_w
            x_row6_right = x_row4_right
            row7_text = str(row.get("row7_text") or "").strip()
            if not row7_text:
                row7_text = _format_lab_row7_pallet_seq(idx, n)
            row7_h = pad + _LAB_ROW7_VALUE_H_MM + pad
            row7_left_w = row4_left_w
            row7_right_w = row4_right_w
            x_row7_right = x_row4_right
            row8_text = str(row.get("row8_text") or "").strip() or _LAB_ROW8_MANUFACTURER_TEXT
            row8_inner_w = content_w - 2 * pad
            row8_text_h = _lab_measure_wrapped_height_mm(
                pdf, row8_text, row8_inner_w, h_txt, bold=False
            )
            row8_h = pad + row8_text_h + pad
            row9_text = (
                str(row.get("row9_text") or "").strip()
                or _LAB_ROW9_PRODUCTION_ADDRESS_TEXT
            )
            row9_inner_w = content_w - 2 * pad
            row9_text_h = _lab_measure_wrapped_height_mm(
                pdf, row9_text, row9_inner_w, h_txt, bold=False
            )
            row9_h = pad + row9_text_h + pad
            row10_batch_lines = _format_lab_row10_batch_lines(
                str(row.get("batch_number") or "")
            )
            row10_expiry = _format_lab_row10_expiry(
                str(row.get("lab_expiry_date") or "")
            )
            row10_batch_block_h = max(h_txt, len(row10_batch_lines) * h_txt)
            row10_h = pad + row10_batch_block_h + pad
            row10_left_w = row4_left_w
            row10_right_w = row4_right_w
            x_row10_right = x_row4_right
            row11_gs1_data, row11_gs1_hri, row11_gs1_err = build_lab_product_gs1_128(
                article,
                str(row.get("batch_number") or ""),
                str(row.get("lab_expiry_date") or ""),
            )
            if row11_gs1_err or not row11_gs1_data:
                return (
                    None,
                    "barcode_fetch",
                    f"Паллета {idx}: {row11_gs1_err or 'нет данных для штрих-кода'}",
                )
            row11_hri_show = str(row.get("row11_gs1_hri") or "").strip() or row11_gs1_hri
            try:
                png_row11 = render_gs1_128_barcode_png(row11_gs1_data, write_text=False)
            except Exception as exc:
                return (
                    None,
                    "barcode_fetch",
                    f"Паллета {idx}: {exc}",
                )
            dims11 = _barcode_raster_pixel_size(png_row11)
            if not dims11:
                return None, "pdf_build", lab_pallet_pdf_error_message("pdf_build")
            pw11, ph11 = dims11
            row11_bc_w = content_w - 2 * pad
            row11_bc_h = row11_bc_w * (ph11 / pw11) * _LAB_BARCODE_HEIGHT_SCALE
            row11_cap_fs = _arnest_barcode_caption_font_pt(pdf, row11_hri_show, row11_bc_w)
            row11_cap_h = _LAB_BARCODE_CAPTION_GAP_MM + _LAB_BARCODE_CAPTION_LINE_H_MM
            row11_h = pad + row11_bc_h + row11_cap_h + pad

            y0 = _LAB_ROW1_TOP_MM
            y_row2 = y0 + row1_h
            y_row3 = y_row2 + row2_h
            y_row4 = y_row3 + row3_h
            y_row5 = y_row4 + row4_h
            y_row6 = y_row5 + row5_h
            y_row7 = y_row6 + row6_h
            y_row8 = y_row7 + row7_h
            y_row9 = y_row8 + row8_h
            y_row10 = y_row9 + row9_h
            y_row11 = y_row10 + row10_h

            pdf.add_page()
            pdf.set_draw_color(0, 0, 0)
            pdf.set_line_width(0.35)
            pdf.rect(x0, y0, left_w, row1_h)
            pdf.rect(x_right, y0, right_w, row1_h)
            pdf.rect(x0, y_row2, content_w, row2_h)
            pdf.rect(x0, y_row3, content_w, row3_h)
            pdf.rect(x0, y_row4, row4_left_w, row4_h)
            pdf.rect(x_row4_right, y_row4, row4_right_w, row4_h)
            pdf.rect(x0, y_row5, row5_left_w, row5_h)
            pdf.rect(x_row5_right, y_row5, row5_right_w, row5_h)
            pdf.rect(x0, y_row6, row6_left_w, row6_h)
            pdf.rect(x_row6_right, y_row6, row6_right_w, row6_h)
            pdf.rect(x0, y_row7, row7_left_w, row7_h)
            pdf.rect(x_row7_right, y_row7, row7_right_w, row7_h)
            pdf.rect(x0, y_row8, content_w, row8_h)
            pdf.rect(x0, y_row9, content_w, row9_h)
            pdf.rect(x0, y_row10, row10_left_w, row10_h)
            pdf.rect(x_row10_right, y_row10, row10_right_w, row10_h)
            pdf.rect(x0, y_row11, content_w, row11_h)

            block_top = y0 + pad
            label_y = block_top + (bc_row_h - h_txt) / 2.0
            bc_y = block_top + (bc_row_h - bc_h) / 2.0

            label_art = "АРТИКУЛ:"
            pdf.set_font("PLCalibri", "B", fs)
            w_art_label = pdf.get_string_width(label_art)
            x_line = x0 + pad
            pdf.set_xy(x_line, label_y)
            pdf.cell(w_art_label, h_txt, label_art, align="L")
            pdf.set_xy(x_line + w_art_label + _LAB_TAB_MM, label_y)
            pdf.cell(w_article, h_txt, art_show, align="L")
            x_label = x_right + pad
            x_bc = x_label + w_pl + _LAB_TAB_MM
            max_bc_x = x0 + content_w - pad - bc_w
            if x_bc > max_bc_x:
                x_bc = max_bc_x
            if x_bc < x_label + w_pl + 2:
                x_bc = x_label + w_pl + _LAB_TAB_MM
            pdf.set_font("PLCalibri", "", fs)
            pdf.set_xy(x_label, label_y)
            pdf.cell(w_pl, h_txt, label_pal)
            pdf.image(
                io.BytesIO(png),
                x=x_bc,
                y=bc_y,
                w=bc_w,
                h=bc_h,
                keep_aspect_ratio=False,
            )
            y_cap = block_top + bc_row_h + _LAB_BARCODE_CAPTION_GAP_MM
            pdf.set_font("PLCalibri", "", cap_fs)
            cap_draw = (
                gs1_hri
                if pdf.get_string_width(gs1_hri) <= bc_w
                else _arnest_clip_text_to_width_mm(pdf, gs1_hri, bc_w)
            )
            pdf.set_xy(x_bc, y_cap)
            pdf.cell(bc_w, _LAB_BARCODE_CAPTION_LINE_H_MM, cap_draw or gs1_hri, align="C")

            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x0 + pad, y_row2 + pad)
            if wm is not None:
                pdf.multi_cell(
                    row2_inner_w,
                    h_txt,
                    product_name,
                    border=0,
                    align="C",
                    wrapmode=wm,
                )
            else:
                pdf.multi_cell(row2_inner_w, h_txt, product_name, border=0, align="C")

            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x0 + pad, y_row3 + pad)
            pdf.cell(row2_inner_w, h_txt, row3_text, align="C")

            row4_label_inner_w = row4_left_w - 2 * pad
            row4_value_inner_w = row4_right_w - 2 * pad
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x0 + pad, y_row4 + pad)
            pdf.cell(row4_label_inner_w, h_txt, _LAB_ROW4_LABEL, align="L")
            pdf.set_xy(x_row4_right + pad, y_row4 + pad)
            pdf.cell(row4_value_inner_w, h_txt, row4_text, align="C")

            row5_label_inner_w = row5_left_w - 2 * pad
            row5_value_inner_w = row5_right_w - 2 * pad
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x0 + pad, y_row5 + pad)
            pdf.cell(row5_label_inner_w, h_txt, _LAB_ROW5_LABEL, align="L")
            pdf.set_font("PLCalibri", "", fs)
            pdf.set_xy(x_row5_right + pad, y_row5 + pad)
            pdf.cell(row5_value_inner_w, h_txt, buyer_order, align="C")

            row6_label_inner_w = row6_left_w - 2 * pad
            row6_value_inner_w = row6_right_w - 2 * pad
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x0 + pad, y_row6 + pad)
            pdf.cell(row6_label_inner_w, h_txt, _LAB_ROW6_LABEL, align="L")
            pdf.set_font("PLCalibri", "", fs)
            pdf.set_xy(x_row6_right + pad, y_row6 + pad)
            pdf.cell(row6_value_inner_w, h_txt, row6_text, align="C")

            row7_label_inner_w = row7_left_w - 2 * pad
            row7_value_inner_w = row7_right_w - 2 * pad
            row7_label_y = y_row7 + pad + (_LAB_ROW7_VALUE_H_MM - h_txt) / 2.0
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x0 + pad, row7_label_y)
            pdf.cell(row7_label_inner_w, h_txt, _LAB_ROW7_LABEL, align="L")
            pdf.set_font("PLCalibri", "B", _LAB_ROW7_VALUE_FONT_PT)
            pdf.set_xy(x_row7_right + pad, y_row7 + pad)
            pdf.cell(
                row7_value_inner_w,
                _LAB_ROW7_VALUE_H_MM,
                row7_text,
                align="C",
            )

            pdf.set_font("PLCalibri", "", fs)
            pdf.set_xy(x0 + pad, y_row8 + pad)
            if wm is not None:
                pdf.multi_cell(
                    row8_inner_w,
                    h_txt,
                    row8_text,
                    border=0,
                    align="L",
                    wrapmode=wm,
                )
            else:
                pdf.multi_cell(row8_inner_w, h_txt, row8_text, border=0, align="L")

            pdf.set_font("PLCalibri", "", fs)
            pdf.set_xy(x0 + pad, y_row9 + pad)
            if wm is not None:
                pdf.multi_cell(
                    row9_inner_w,
                    h_txt,
                    row9_text,
                    border=0,
                    align="L",
                    wrapmode=wm,
                )
            else:
                pdf.multi_cell(row9_inner_w, h_txt, row9_text, border=0, align="L")

            row10_left_inner = row10_left_w - 2 * pad
            row10_right_inner = row10_right_w - 2 * pad
            y_row10_inner = y_row10 + pad
            _lab_draw_row10_batch_block(
                pdf,
                x0 + pad,
                y_row10_inner,
                row10_left_inner,
                row10_batch_lines,
                h_txt,
                fs,
                _LAB_TAB_MM,
            )
            row10_expiry_y = y_row10 + pad + (row10_batch_block_h - h_txt) / 2.0
            _lab_draw_bold_label_tab_value(
                pdf,
                x_row10_right + pad,
                row10_expiry_y,
                row10_right_inner,
                _LAB_ROW10_EXPIRY_LABEL,
                row10_expiry,
                h_txt,
                fs,
                _LAB_TAB_MM,
            )

            row11_bc_x = x0 + pad
            row11_bc_y = y_row11 + pad
            pdf.image(
                io.BytesIO(png_row11),
                x=row11_bc_x,
                y=row11_bc_y,
                w=row11_bc_w,
                h=row11_bc_h,
                keep_aspect_ratio=False,
            )
            y_row11_cap = row11_bc_y + row11_bc_h + _LAB_BARCODE_CAPTION_GAP_MM
            pdf.set_font("PLCalibri", "", row11_cap_fs)
            row11_cap_draw = (
                row11_hri_show
                if pdf.get_string_width(row11_hri_show) <= row11_bc_w
                else _arnest_clip_text_to_width_mm(pdf, row11_hri_show, row11_bc_w)
            )
            pdf.set_xy(row11_bc_x, y_row11_cap)
            pdf.cell(
                row11_bc_w,
                _LAB_BARCODE_CAPTION_LINE_H_MM,
                row11_cap_draw or row11_hri_show,
                align="C",
            )

        out = pdf.output()
        return (bytes(out) if isinstance(out, (bytes, bytearray)) else out.encode("latin-1")), None, None
    except Exception:
        return None, "pdf_build", lab_pallet_pdf_error_message("pdf_build")


def build_lab_industries_pallet_sheets_pdf_from_order(
    detail: dict[str, Any],
    *,
    lab_sscc_ai_seq: int | None = None,
) -> tuple[bytes | None, str | None, str | None]:
    if not is_lab_industries_client(str(detail.get("client") or "")):
        return None, "lab_client", lab_pallet_pdf_error_message("lab_client")
    if lab_sscc_ai_seq is None:
        lab_sscc_ai_seq = int(detail.get("lab_sscc_ai_seq") or 0) % 100
    detail_with_seq = {**detail, "lab_sscc_ai_seq": lab_sscc_ai_seq}
    pallets = lab_pallets_from_order_detail(detail_with_seq)
    if not pallets:
        st = detail.get("assemble_state")
        if not isinstance(st, dict):
            return None, "no_assembly", lab_pallet_pdf_error_message("no_assembly")
        return None, "no_pallets", lab_pallet_pdf_error_message("no_pallets")
    return build_lab_industries_pallet_sheets_pdf_bytes(pallets)

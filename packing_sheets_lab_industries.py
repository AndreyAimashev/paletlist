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


def build_lab_pallet_sscc_gs1(pallet_number: str, fallback_index: int = 1) -> tuple[str, str]:
    """GS1-128 (AI 00): код для генератора и подпись (00)150000000000000001."""
    n = _parse_pallet_number_int(pallet_number, fallback_index)
    sscc18 = f"150000000000000{n:03d}"
    if len(sscc18) != 18:
        sscc18 = (sscc18 + "0" * 18)[:18]
    encode = "00" + sscc18
    human = f"(00){sscc18}"
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


def _format_lab_row3_text(total_pieces: float, volume_ml: float) -> str:
    """Например: 2394 х 200мл."""
    pcs = max(0, int(round(total_pieces)))
    vol = float(volume_ml or 0)
    if vol <= 0:
        return str(pcs)
    if abs(vol - round(vol)) < 1e-6:
        vol_s = str(int(round(vol)))
    else:
        vol_s = ("%g" % vol).replace(".", ",")
    return f"{pcs} х {vol_s}мл"


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
    out: list[dict[str, Any]] = []
    for idx, pal in enumerate(pallets, start=1):
        pnum = str(pal.get("palletNumber") or "").strip()
        if not pnum:
            pnum = str(idx)
        line_idx, article, name = _first_line_on_pallet(pal, items)
        total_pcs = totals.get(line_idx, 0.0) if line_idx is not None else 0.0
        vol = 0.0
        if line_idx is not None and 0 <= line_idx < len(items):
            vol = float(items[line_idx].get("volume_ml") or 0)
        out.append(
            {
                "article": article,
                "name": name,
                "pallet_number": pnum,
                "total_pieces": total_pcs,
                "volume_ml": vol,
                "row3_text": _format_lab_row3_text(total_pcs, vol),
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
    """PDF: строки 1–3 (артикул / паллета GS1-128 / наименование / кол-во и объём)."""
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
                row3_text = _format_lab_row3_text(total_pcs, vol)
            pallet_number = str(row.get("pallet_number") or "").strip()
            gs1_data, gs1_hri = build_lab_pallet_sscc_gs1(pallet_number, idx)
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
            bc_h = bc_w * (ph / pw) * _LAB_BARCODE_HEIGHT_SCALE
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

            y0 = _LAB_ROW1_TOP_MM
            y_row2 = y0 + row1_h
            y_row3 = y_row2 + row2_h

            pdf.add_page()
            pdf.set_draw_color(0, 0, 0)
            pdf.set_line_width(0.35)
            pdf.rect(x0, y0, left_w, row1_h)
            pdf.rect(x_right, y0, right_w, row1_h)
            pdf.rect(x0, y_row2, content_w, row2_h)
            pdf.rect(x0, y_row3, content_w, row3_h)

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

        out = pdf.output()
        return (bytes(out) if isinstance(out, (bytes, bytearray)) else out.encode("latin-1")), None, None
    except Exception:
        return None, "pdf_build", lab_pallet_pdf_error_message("pdf_build")


def build_lab_industries_pallet_sheets_pdf_from_order(
    detail: dict[str, Any],
) -> tuple[bytes | None, str | None, str | None]:
    if not is_lab_industries_client(str(detail.get("client") or "")):
        return None, "lab_client", lab_pallet_pdf_error_message("lab_client")
    pallets = lab_pallets_from_order_detail(detail)
    if not pallets:
        st = detail.get("assemble_state")
        if not isinstance(st, dict):
            return None, "no_assembly", lab_pallet_pdf_error_message("no_assembly")
        return None, "no_pallets", lab_pallet_pdf_error_message("no_pallets")
    return build_lab_industries_pallet_sheets_pdf_bytes(pallets)

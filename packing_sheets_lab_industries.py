# -*- coding: utf-8 -*-
"""Паллетные листы для клиента «ЛАБ Индастриз» (PDF, GS1-128 SSCC для номера паллеты)."""

from __future__ import annotations

import io
import re
from typing import Any

_LAB_CLIENT_NORMALIZED = "лаб индастриз"

# A4 портрет, поля как у паллетных листов Арнест.
_LAB_A4_W_MM = 210.0
_LAB_MARGIN_MM = 12.0
_LAB_BOX_GAP_MM = 6.0
_LAB_BOX_PAD_MM = 3.0
_LAB_ROW1_TOP_MM = 14.0
_LAB_TEXT_FONT_PT = 12.0
_LAB_TEXT_H_MM = 7.0
_LAB_TAB_MM = 12.5
_LAB_BARCODE_W_MM = 72.0
_LAB_BARCODE_HEIGHT_SCALE = 0.55
_LAB_BARCODE_CAPTION_GAP_MM = 1.5
_LAB_BARCODE_CAPTION_LINE_H_MM = 6.0
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


def _first_article_on_pallet(pal: dict[str, Any], items: list[dict[str, Any]]) -> str:
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
            art = str(items[idx].get("article") or "").strip()
            if art:
                return art
    return ""


def lab_pallets_from_order_detail(detail: dict[str, Any]) -> list[dict[str, str]]:
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
    out: list[dict[str, str]] = []
    for idx, pal in enumerate(pallets, start=1):
        pnum = str(pal.get("palletNumber") or "").strip()
        if not pnum:
            pnum = str(idx)
        out.append(
            {
                "article": _first_article_on_pallet(pal, items) or "—",
                "pallet_number": pnum,
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


def build_lab_industries_pallet_sheets_pdf_bytes(
    pallets: list[dict[str, Any]],
) -> tuple[bytes | None, str | None, str | None]:
    """PDF: пока реализована строка 1 (артикул слева, номер паллеты GS1-128 справа)."""
    from api_server import (  # noqa: PLC0415 — избегаем циклического импорта на уровне модуля
        Align,
        FPDF,
        HAVE_FPDF,
        _arnest_barcode_caption_font_pt,
        _arnest_clip_text_to_width_mm,
        _arnest_pdf_register_text_fonts,
        _barcode_raster_pixel_size,
        render_gs1_128_barcode_png,
    )

    if not HAVE_FPDF or FPDF is None or Align is None:
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
        box_w = (content_w - _LAB_BOX_GAP_MM) / 2.0
        fs = _LAB_TEXT_FONT_PT
        h_txt = _LAB_TEXT_H_MM
        pad = _LAB_BOX_PAD_MM
        for idx, row in enumerate(rows, start=1):
            article = str(row.get("article") or "").strip() or "—"
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
            bc_h = _LAB_BARCODE_W_MM * (ph / pw) * _LAB_BARCODE_HEIGHT_SCALE
            cap_fs = _arnest_barcode_caption_font_pt(pdf, gs1_hri, box_w - 2 * pad)
            cap_h = _LAB_BARCODE_CAPTION_GAP_MM + _LAB_BARCODE_CAPTION_LINE_H_MM
            box_h = max(
                pad + h_txt + pad,
                pad + h_txt + _LAB_BARCODE_CAPTION_GAP_MM + bc_h + cap_h + pad,
            )
            y0 = _LAB_ROW1_TOP_MM
            x_left = _LAB_MARGIN_MM
            x_right = _LAB_MARGIN_MM + box_w + _LAB_BOX_GAP_MM
            pdf.add_page()
            pdf.set_draw_color(0, 0, 0)
            pdf.set_line_width(0.35)
            pdf.rect(x_left, y0, box_w, box_h)
            pdf.rect(x_right, y0, box_w, box_h)
            label_art = "АРТИКУЛ:"
            pdf.set_font("PLCalibri", "B", fs)
            w_label = pdf.get_string_width(label_art)
            x_art = x_left + pad
            y_art = y0 + pad
            pdf.set_xy(x_art, y_art)
            pdf.cell(w_label, h_txt, label_art)
            art_max = max(5.0, box_w - 2 * pad - w_label - _LAB_TAB_MM)
            art_show = _arnest_clip_text_to_width_mm(pdf, article, art_max)
            pdf.set_xy(x_art + w_label + _LAB_TAB_MM, y_art)
            pdf.cell(art_max, h_txt, art_show or "—")
            label_pal = "Номер паллета:"
            pdf.set_font("PLCalibri", "", fs)
            w_pl = pdf.get_string_width(label_pal)
            x_pl = x_right + pad
            y_pl = y0 + pad
            pdf.set_xy(x_pl, y_pl)
            pdf.cell(box_w - 2 * pad, h_txt, label_pal)
            bc_x = x_right + (box_w - _LAB_BARCODE_W_MM) / 2.0
            bc_y = y_pl + h_txt + _LAB_BARCODE_CAPTION_GAP_MM
            pdf.image(
                io.BytesIO(png),
                x=bc_x,
                y=bc_y,
                w=_LAB_BARCODE_W_MM,
                h=bc_h,
                keep_aspect_ratio=False,
            )
            y_cap = bc_y + bc_h + _LAB_BARCODE_CAPTION_GAP_MM
            pdf.set_font("PLCalibri", "", cap_fs)
            cap_draw = (
                gs1_hri
                if pdf.get_string_width(gs1_hri) <= box_w - 2 * pad
                else _arnest_clip_text_to_width_mm(pdf, gs1_hri, box_w - 2 * pad)
            )
            pdf.set_xy(x_right + pad, y_cap)
            pdf.cell(box_w - 2 * pad, _LAB_BARCODE_CAPTION_LINE_H_MM, cap_draw or gs1_hri, align="C")
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

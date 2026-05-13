#!/usr/bin/env python3
import datetime
import io
import json
import math
import os
import re
import socket
import sqlite3
import threading
import unicodedata
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

try:
    from fpdf import FPDF
    from fpdf.enums import Align, WrapMode

    HAVE_FPDF = True
except ImportError:
    HAVE_FPDF = False
    FPDF = None  # type: ignore
    Align = None  # type: ignore
    WrapMode = None  # type: ignore

try:
    from barcode import Code128
    from barcode.writer import ImageWriter

    HAVE_CODE128_BARCODE = True
except ImportError:
    HAVE_CODE128_BARCODE = False
    Code128 = None  # type: ignore
    ImageWriter = None  # type: ignore

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False
    Workbook = None  # type: ignore
    Alignment = None  # type: ignore
    Border = None  # type: ignore
    Font = None  # type: ignore
    PatternFill = None  # type: ignore
    Side = None  # type: ignore
    get_column_letter = None  # type: ignore


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "warehouse.db"
JSON_SEED_PATH = BASE_DIR / "nomenclature.json"
HOST = "127.0.0.1"

try:
    from packing_sheets_generic import (
        build_generic_packing_sheets_html,
        build_generic_packing_sheets_pdf_bytes,
        generic_packing_error_message,
    )
except ImportError:
    build_generic_packing_sheets_html = None  # type: ignore[misc, assignment]
    build_generic_packing_sheets_pdf_bytes = None  # type: ignore[misc, assignment]
    generic_packing_error_message = None  # type: ignore[misc, assignment]

PORT = 8081
DB_LOCK = threading.Lock()


def _norm_api_path(path: str) -> str:
    p = (path or "").strip()
    while "//" in p:
        p = p.replace("//", "/")
    p = p.rstrip("/")
    return p if p else "/"


def _is_nomenclature_list_path(path: str) -> bool:
    return _norm_api_path(path) == "/api/nomenclature"


def _is_orders_list_path(path: str) -> bool:
    return _norm_api_path(path) == "/api/orders"


def _is_arnest_unirus_pallet_sheets_pdf_path(path: str) -> bool:
    return _norm_api_path(path) == "/api/arnest-unirus-pallet-sheets-pdf"


def _is_arnest_unirus_pallet_sheets_xlsx_path(path: str) -> bool:
    return _norm_api_path(path) == "/api/arnest-unirus-pallet-sheets-xlsx"


def _is_print_arnest_unirus_pallet_sheets_raw_path(path: str) -> bool:
    return _norm_api_path(path) == "/api/print-arnest-unirus-pallet-sheets-raw"


def _is_pallet_printer_raw_ping_path(path: str) -> bool:
    return _norm_api_path(path) == "/api/pallet-printer-raw-ping"


# Префикс Code128 данных паллеты (как в макросе Word / tec-it).
PALLET_BARCODE_PREFIX = "1500000"
MAX_PALLET_SHEET_PDF_PAGES = 500


def _format_pallet_suffix_vba(n: int) -> str:
    """Эквивалент VBA Format(n, \"000\"): минимум три знака с ведущими нулями."""
    if n < 0:
        n = -n
    s = str(n)
    if len(s) < 3:
        return s.zfill(3)
    return s


def build_pallet_barcode_data(pallet_number: str) -> str | None:
    """Строка данных Code128: 1500000 + суффикс номера паллеты; None если номер пустой."""
    raw = str(pallet_number).strip()
    if not raw:
        return None
    try:
        n = int(float(raw.replace(",", ".")))
    except ValueError:
        digits = re.sub(r"\D", "", raw)
        if not digits:
            return None
        try:
            n = int(digits)
        except ValueError:
            return None
    return PALLET_BARCODE_PREFIX + _format_pallet_suffix_vba(n)


def render_code128_barcode_png(barcode_data: str, *, write_text: bool = True) -> bytes:
    """Code128 в PNG (python-barcode + Pillow). При write_text=False только полосы — подпись в PDF отдельно."""
    if not HAVE_CODE128_BARCODE or Code128 is None or ImageWriter is None:
        raise RuntimeError(
            "Не установлен python-barcode[images] (pip install \"python-barcode[images]\")"
        )
    payload = str(barcode_data)
    if not payload.strip():
        raise ValueError("Пустые данные для штрих-кода")
    writer = ImageWriter(format="PNG", dpi=300)
    buf = io.BytesIO()
    Code128(payload, writer=writer).write(buf, options={"write_text": write_text})
    out = buf.getvalue()
    if not out or _png_ihdr_pixel_size(out) is None:
        raise RuntimeError("Пустой или некорректный PNG штрих-кода")
    return out


def _png_ihdr_pixel_size(png: bytes) -> tuple[int, int] | None:
    """Ширина и высота растра из PNG IHDR (первый чанк)."""
    if len(png) < 24 or png[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    if png[12:16] != b"IHDR":
        return None
    w = int.from_bytes(png[16:20], "big")
    h = int.from_bytes(png[20:24], "big")
    if w <= 0 or h <= 0:
        return None
    return w, h


def _gif_logical_screen_size(gif: bytes) -> tuple[int, int] | None:
    """Логический экран GIF87a/GIF89a (ширина и высота в пикселях)."""
    if len(gif) < 10 or gif[:6] not in (b"GIF87a", b"GIF89a"):
        return None
    w = int.from_bytes(gif[6:8], "little")
    h = int.from_bytes(gif[8:10], "little")
    if w <= 0 or h <= 0:
        return None
    return w, h


def _barcode_raster_pixel_size(raw: bytes) -> tuple[int, int] | None:
    """Размеры растра штрих-кода (PNG локальной генерации или GIF)."""
    z = _png_ihdr_pixel_size(raw)
    if z:
        return z
    return _gif_logical_screen_size(raw)


# A4 и поля как в типовом документе Word (≈2 см слева/справа).
_ARNEST_A4_W_MM = 210.0
_ARNEST_SIDE_MARGIN_MM = 20.0
# Строка 1: высота растра полос из соотношения сторон её PNG и _ARNEST_BARCODE_HEIGHT_SCALE.
# Строка 8: ширина _ARNEST_LINE8_BARCODE_W_MM, высота из PNG × scale.
# Строка 14 (паллета): та же ширина; высота принудительно как у строки 8, чтобы полосы были одинаковой высоты.
_ARNEST_BARCODE_HEIGHT_SCALE = 0.5
_ARNEST_BARCODE_CAPTION_GAP_MM = 1.5  # между низом растра полос и подписью (PDF)
_ARNEST_BARCODE_CAPTION_LINE_H_MM = 6.0
_ARNEST_BARCODE_CAPTION_FONT_MAX_PT = 11.0
_ARNEST_BARCODE_CAPTION_FONT_MIN_PT = 7.0
_ARNEST_LINE2_GAP_MM = 10.0  # только между строкой 1 (штрих-код) и строкой 2
_ARNEST_TEXT_LINE_GAP_MM = 5.0  # между всеми остальными строками (текст и нижний штрих-код)
_ARNEST_LINE4_NAME_EXTRA_GAP_MM = 4.0  # доп. зазор после блока строки 4 (наименование может занять 2+ строки)
_ARNEST_TEXT_FONT_PT = 12.0  # все текстовые элементы паллетного листа (не штрих-код)
_ARNEST_LINE2_TEXT_H_MM = 7.0
_ARNEST_LINE2_ARTICLE_MAX_W_MM = 105.0
_ARNEST_TAB_MM = 12.5  # одна стандартная табуляция (~1,25 см) между артикулом и наименованием в строке 4
_ARNEST_ASEPTICA_LABEL = "ASEPTICA"
_ARNEST_LINE3_DESCRIPTION_LABEL = "Description"
_ARNEST_LINE5_GTIN_LABEL = "GTIN CODE"
_ARNEST_LINE6_MGT_DATE_LABEL = "Mgt. Date"
_ARNEST_LINE6_EXPIRY_DATE_LABEL = "Expiry Date"
_ARNEST_LINE6_CROSS_WEIGHT_LABEL = "Cross Weight(KG)"
_ARNEST_LINE8_BARCODE_W_MM = 45.0
_ARNEST_LINE9_PALLET_QTY_LABEL = "Pallet Quantity"
_ARNEST_LINE9_UNITS_PC_LABEL = "Units PC"
_ARNEST_LINE9_TI_HI_LABEL = "TI*HI"
_ARNEST_LINE10_KOR_LABEL = "кор."
_ARNEST_LINE12_TOTAL_QUANTITY_LABEL = "Total Quantity"
_ARNEST_LINE15_BATCH_NUMBER_LABEL = "Batch Number"


def _arnest_regular_font_path() -> Path | None:
    """Обычный Calibri / запасной DejaVu Sans / Liberation для наименования в строке 4."""
    windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
    win_fonts = windir / "Fonts"
    candidates = [
        BASE_DIR / "fonts" / "calibri.ttf",
        BASE_DIR / "fonts" / "Calibri.ttf",
        win_fonts / "calibri.ttf",
        win_fonts / "CALIBRI.TTF",
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/TTF/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
    ]
    for c in candidates:
        if c.is_file():
            return c
    return None


def _arnest_line2_bold_font_path() -> Path | None:
    """Жирный TTF: Calibri Bold при наличии, иначе типичный системный запасной (DejaVu на Linux)."""
    windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
    win_fonts = windir / "Fonts"
    candidates = [
        BASE_DIR / "fonts" / "calibrib.ttf",
        BASE_DIR / "fonts" / "Calibrib.ttf",
        win_fonts / "calibrib.ttf",
        win_fonts / "CALIBRIB.TTF",
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        Path("/usr/share/fonts/TTF/DejaVuSans-Bold.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
    ]
    for c in candidates:
        if c.is_file():
            return c
    return None


def _arnest_pdf_register_text_fonts(pdf: FPDF) -> str | None:
    """Семейство PLCalibri: обычный (наименование) и жирный (артикул, ASEPTICA, Description)."""
    bold_p = _arnest_line2_bold_font_path()
    if bold_p is None:
        return "no_line2_font"
    reg_p = _arnest_regular_font_path()
    try:
        pdf.add_font("PLCalibri", "B", str(bold_p))
        if reg_p is not None:
            pdf.add_font("PLCalibri", "", str(reg_p))
        else:
            pdf.add_font("PLCalibri", "", str(bold_p))
    except OSError:
        return "no_line2_font"
    return None


def _arnest_clip_text_to_width_mm(pdf: FPDF, text: str, max_mm: float) -> str:
    """Подрезает строку по ширине (мм) с многоточием, шрифт уже выбран."""
    t = text or ""
    if pdf.get_string_width(t) <= max_mm:
        return t
    ell = "…"
    if pdf.get_string_width(ell) > max_mm:
        return ""
    n = len(t)
    while n > 0:
        cand = t[:n] + ell
        if pdf.get_string_width(cand) <= max_mm:
            return cand
        n -= 1
    return ell


def _arnest_barcode_caption_font_pt(pdf: FPDF, text: str, max_w_mm: float) -> float:
    """Размер шрифта подписи под Code128: уменьшаем, пока строка влезает в max_w_mm."""
    t = str(text) if text is not None else ""
    lo = _ARNEST_BARCODE_CAPTION_FONT_MIN_PT
    hi = _ARNEST_BARCODE_CAPTION_FONT_MAX_PT
    pt = hi
    step = 0.5
    while pt >= lo:
        pdf.set_font("PLCalibri", "", pt)
        if pdf.get_string_width(t) <= max_w_mm:
            return pt
        pt -= step
    pdf.set_font("PLCalibri", "", lo)
    return lo


def _arnest_draw_code128_caption_below(
    pdf: FPDF,
    x: float,
    w_mm: float,
    y_barcode_bottom: float,
    text: str,
    *,
    body_caption_style: bool = False,
) -> float:
    """Подпись под штрих-кодом (отдельно от PNG). Возвращает занятую высоту: зазор + строка.

    body_caption_style: как цифры на листе (партия и т.д.) — 12 pt жирный, без уменьшения кегля.
    """
    gap = _ARNEST_BARCODE_CAPTION_GAP_MM
    s = str(text) if text is not None else ""
    if body_caption_style:
        h_line = _ARNEST_LINE2_TEXT_H_MM
        fs = _ARNEST_TEXT_FONT_PT
        pdf.set_font("PLCalibri", "B", fs)
        draw = _arnest_clip_text_to_width_mm(pdf, s, w_mm) if s else ""
        pdf.set_xy(x, y_barcode_bottom + gap)
        pdf.cell(w_mm, h_line, draw or s, align="C")
        return gap + h_line
    h_line = _ARNEST_BARCODE_CAPTION_LINE_H_MM
    fs = _arnest_barcode_caption_font_pt(pdf, text, w_mm)
    pdf.set_font("PLCalibri", "", fs)
    draw = s if pdf.get_string_width(s) <= w_mm else _arnest_clip_text_to_width_mm(pdf, s, w_mm)
    pdf.set_xy(x, y_barcode_bottom + gap)
    pdf.cell(w_mm, h_line, draw or s, align="C")
    return gap + h_line


def _arnest_pallet_pdf_error_message(code: str) -> str:
    return {
        "no_fpdf": "На сервере не установлен пакет fpdf2 (pip install fpdf2).",
        "validation": "Укажите число страниц от 1 до 500 (по одной на паллету).",
        "validation_pallets": "Передайте непустой массив pallets (не более 500 паллет).",
        "pdf_build": "Не удалось сформировать PDF.",
        "no_barcode": (
            "На сервере не установлен python-barcode[images] "
            '(pip install "python-barcode[images]").'
        ),
        "barcode_fetch": "Не удалось сформировать изображение штрих-кода.",
        "no_line2_font": (
            "Не найдены TTF для текста паллетного PDF (жирный обязателен). На сервере: apt install fonts-dejavu-core "
            "или положите calibri.ttf и calibrib.ttf в каталог fonts/ рядом с api_server.py."
        ),
        "no_openpyxl": 'На сервере не установлен openpyxl (pip install openpyxl).',
        "xlsx_build": "Не удалось сформировать Excel-файл.",
    }.get(code, "Ошибка генерации PDF.")


def _arnest_first_digit_run(article: str) -> str | None:
    m = re.search(r"\d+", article or "")
    return m.group(0) if m else None


def _arnest_yymmdd_digits(raw: str) -> str | None:
    d = re.sub(r"\D", "", raw or "")
    if len(d) >= 6:
        d = d[-6:]
    if len(d) != 6 or not d.isdigit():
        return None
    return d


def _arnest_yymmdd_for_barcode(raw: str) -> str | None:
    """Шесть цифр из поля → в штрих-коде пары AB·CD·EF дают EF·CD·AB (напр. 123456 → 563412)."""
    d = _arnest_yymmdd_digits(raw)
    if not d:
        return None
    return d[4:6] + d[2:4] + d[0:2]


def _arnest_six_digits_display_ddmmyy(raw: str) -> str | None:
    """Шесть цифр → «ДД.ММ.ГГГГ» для строки 7: сначала ДДММГГ, при невалидной дате — ГГММДД.

    Если оба разбора невалидны (напр. несуществующий месяц вроде «22»), возвращается None —
    в PDF подставляется «—»; выдумывать дату не нужно.
    """
    d = re.sub(r"\D", "", raw or "")
    if len(d) >= 6:
        d = d[-6:]
    if len(d) != 6 or not d.isdigit():
        return None

    dd = int(d[0:2])
    mm = int(d[2:4])
    yy = int(d[4:6])
    year = 2000 + yy
    try:
        datetime.date(year, mm, dd)
        return f"{dd:02d}.{mm:02d}.{year:04d}"
    except ValueError:
        pass

    yy2 = int(d[0:2])
    mm2 = int(d[2:4])
    dd2 = int(d[4:6])
    yfull = 2000 + yy2
    try:
        datetime.date(yfull, mm2, dd2)
        return f"{dd2:02d}.{mm2:02d}.{yfull:04d}"
    except ValueError:
        return None


def _arnest_coerce_date_string(val) -> str | None:
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        if isinstance(val, float):
            if not math.isfinite(val):
                return None
            if abs(val - round(val)) < 1e-9:
                s = str(int(round(val)))
            else:
                s = str(val).strip()
        else:
            s = str(val)
    else:
        s = str(val).strip()
    if not s or s.lower() == "none":
        return None
    return s


def _arnest_unirus_date_raw(row: dict, *keys: str) -> str:
    """Дата из полей верхнего уровня паллеты (учёт JSON null)."""
    for key in keys:
        if key not in row:
            continue
        s = _arnest_coerce_date_string(row[key])
        if s:
            return s
    return ""


def _arnest_date_raw_from_slots(row: dict, *slot_keys: str) -> str:
    """Дата из slots[] (если в объекте паллеты не продублированы наверх)."""
    slots = row.get("slots")
    if not isinstance(slots, list):
        return ""
    for slot in slots:
        if not isinstance(slot, dict):
            continue
        for sk in slot_keys:
            if sk not in slot:
                continue
            s = _arnest_coerce_date_string(slot.get(sk))
            if s:
                return s
    return ""


def _arnest_mfg_date_raw(row: dict) -> str:
    s = _arnest_unirus_date_raw(
        row,
        "manufacturing_date_raw",
        "manufacturingDateRaw",
        "unirusMfgDate",
    )
    if s:
        return s
    return _arnest_date_raw_from_slots(
        row, "unirusMfgDate", "manufacturingDateRaw", "manufacturing_date_raw"
    )


def _arnest_exp_date_raw(row: dict) -> str:
    s = _arnest_unirus_date_raw(
        row,
        "expiry_date_raw",
        "expiryDateRaw",
        "unirusExpiryDate",
    )
    if s:
        return s
    return _arnest_date_raw_from_slots(
        row, "unirusExpiryDate", "expiryDateRaw", "expiry_date_raw"
    )


def _arnest_format_weight_kg_ru(raw) -> str:
    """Вес в кг для PDF: одна десятичная, запятая как десятичный разделитель."""
    try:
        v = float(raw)
    except (TypeError, ValueError):
        v = 0.0
    if not math.isfinite(v) or v < 0:
        v = 0.0
    s = f"{v:.1f}"
    return s.replace(".", ",")


def _arnest_format_boxes_qty_for_barcode(raw) -> str:
    """Количество коробок для штрих-кода строки 8."""
    try:
        v = float(raw)
    except (TypeError, ValueError):
        v = 0.0
    if not math.isfinite(v) or v < 0:
        v = 0.0
    if abs(v - round(v)) < 1e-9:
        return str(int(round(v)))
    return f"{v:.1f}".rstrip("0").rstrip(".")


def _arnest_remainder_boxes_star_one(boxes_f: float, rl: int) -> str:
    """Остаток коробок после полных рядов: всего − (ряд × полные ряды), затем «*1» (напр. 5*1)."""
    if rl <= 0 or not math.isfinite(boxes_f) or boxes_f < 0:
        return "—*1"
    full_rows = int(math.floor(boxes_f / rl + 1e-9))
    rem = max(0.0, boxes_f - rl * full_rows)
    if abs(rem - round(rem)) < 1e-9:
        rem_s = str(int(round(rem)))
    else:
        rem_s = f"{rem:.1f}".rstrip("0").rstrip(".")
    return f"{rem_s}*1"


def _arnest_line_barcode_data(row: dict) -> tuple[str | None, str | None]:
    """Строка Code128: артикул (цифры), партия, дата изготовления и срок в кодировке для сканера (пары цифр дат переставлены)."""
    article = str(row.get("article", "")).strip()
    part_art = _arnest_first_digit_run(article)
    if not part_art:
        return None, "в артикуле нет цифр"
    batch = str(row.get("batch_number", row.get("batchNumber", ""))).strip()
    if not batch:
        return None, "не указана партия"
    mfg = _arnest_yymmdd_for_barcode(_arnest_mfg_date_raw(row))
    if not mfg:
        return None, "дата изготовления: нужны 6 цифр ДДММГГ"
    exp = _arnest_yymmdd_for_barcode(_arnest_exp_date_raw(row))
    if not exp:
        return None, "срок годности: нужны 6 цифр ДДММГГ"
    return f"{part_art} {batch} {mfg} {exp}", None


def build_arnest_unirus_pallet_sheets_pdf_with_barcodes(
    pallets: list[dict],
) -> tuple[bytes | None, str, str]:
    """Один PDF: 1, 8, 14 — Code128 (только полосы в PNG, читаемая строка под ними в PDF); 2–16 — текст 12 pt; 14 — паллета слева (если номер); 15–16 — партия у колонки «кор.»."""
    if not HAVE_FPDF or FPDF is None or Align is None:
        return None, "no_fpdf", ""
    if not HAVE_CODE128_BARCODE:
        return None, "no_barcode", _arnest_pallet_pdf_error_message("no_barcode")
    n = len(pallets)
    if n < 1 or n > MAX_PALLET_SHEET_PDF_PAGES:
        return None, "validation_pallets", ""
    try:
        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.set_auto_page_break(False)
        font_err = _arnest_pdf_register_text_fonts(pdf)
        if font_err:
            return None, font_err, _arnest_pallet_pdf_error_message(font_err)
        barcode_w_mm = _ARNEST_A4_W_MM - 2 * _ARNEST_SIDE_MARGIN_MM
        y_top = 14.0
        for idx, row in enumerate(pallets, start=1):
            if not isinstance(row, dict):
                return None, "validation_pallet", f"Паллета {idx}: ожидался объект с полями."
            data, err_detail = _arnest_line_barcode_data(row)
            if err_detail:
                return None, "validation_pallet", f"Паллета {idx}: {err_detail}."
            try:
                png = render_code128_barcode_png(data, write_text=False)
            except Exception as exc:
                return (
                    None,
                    "barcode_fetch",
                    f"Не удалось сформировать штрих-код для паллеты {idx}: {exc}",
                )
            dims = _barcode_raster_pixel_size(png)
            if not dims:
                return None, "pdf_build", ""
            pw, ph = dims
            h_base_mm = barcode_w_mm * (ph / pw)
            h_mm = h_base_mm * _ARNEST_BARCODE_HEIGHT_SCALE
            pdf.add_page()
            pdf.image(
                io.BytesIO(png),
                x=Align.C,
                y=y_top,
                w=barcode_w_mm,
                h=h_mm,
                keep_aspect_ratio=False,
            )
            cap1_h = _arnest_draw_code128_caption_below(
                pdf, _ARNEST_SIDE_MARGIN_MM, barcode_w_mm, y_top + h_mm, data
            )
            y2 = y_top + h_mm + cap1_h + _ARNEST_LINE2_GAP_MM
            fs = _ARNEST_TEXT_FONT_PT
            h_txt = _ARNEST_LINE2_TEXT_H_MM
            left_w = _ARNEST_LINE2_ARTICLE_MAX_W_MM
            art_full = str(row.get("article", "")).strip()
            pdf.set_font("PLCalibri", "BU", fs)
            art_show = (
                _arnest_clip_text_to_width_mm(pdf, art_full, left_w) if art_full else "—"
            )
            if not art_show:
                art_show = "—"
            pdf.set_xy(_ARNEST_SIDE_MARGIN_MM, y2)
            pdf.cell(left_w, h_txt, art_show)
            pdf.set_font("PLCalibri", "B", fs)
            ase = _ARNEST_ASEPTICA_LABEL
            w_ase = pdf.get_string_width(ase)
            pdf.set_xy((_ARNEST_A4_W_MM - w_ase) / 2.0, y2)
            pdf.cell(w_ase, h_txt, ase)
            y3 = y2 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            content_w = _ARNEST_A4_W_MM - 2 * _ARNEST_SIDE_MARGIN_MM
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(_ARNEST_SIDE_MARGIN_MM, y3)
            pdf.cell(content_w, h_txt, _ARNEST_LINE3_DESCRIPTION_LABEL, align="L")
            y4 = y3 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            x0 = _ARNEST_SIDE_MARGIN_MM
            max_r = _ARNEST_A4_W_MM - _ARNEST_SIDE_MARGIN_MM
            art_l4 = str(row.get("article", "")).strip() or "—"
            name_raw = str(row.get("name", row.get("product_name", ""))).strip() or "—"
            pdf.set_font("PLCalibri", "B", fs)
            w_art0 = pdf.get_string_width(art_l4)
            max_art = max(10.0, max_r - x0 - _ARNEST_TAB_MM - 25.0)
            art_l4_draw = (
                _arnest_clip_text_to_width_mm(pdf, art_l4, max_art)
                if w_art0 > max_art
                else art_l4
            )
            w_art = pdf.get_string_width(art_l4_draw)
            pdf.set_xy(x0, y4)
            pdf.cell(w_art, h_txt, art_l4_draw)
            x_name = x0 + w_art + _ARNEST_TAB_MM
            pdf.set_font("PLCalibri", "", fs)
            rem_w = max(0.0, max_r - x_name)
            name_txt = name_raw if name_raw else "—"
            name_block_h = h_txt
            if rem_w > 0.5:
                pdf.set_xy(x_name, y4)
                wm = WrapMode.WORD if WrapMode is not None else None
                if wm is not None:
                    pdf.multi_cell(
                        rem_w,
                        h_txt,
                        name_txt,
                        border=0,
                        align="L",
                        wrapmode=wm,
                    )
                else:
                    pdf.multi_cell(rem_w, h_txt, name_txt, border=0, align="L")
                name_block_h = max(h_txt, pdf.get_y() - y4)
            else:
                pdf.set_xy(x_name, y4)
                pdf.cell(1.0, h_txt, _arnest_clip_text_to_width_mm(pdf, name_txt, 1.0) or "—")
            y5 = (
                y4
                + max(h_txt, name_block_h)
                + _ARNEST_LINE4_NAME_EXTRA_GAP_MM
                + _ARNEST_TEXT_LINE_GAP_MM
            )
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x0, y5)
            pdf.cell(content_w, h_txt, _ARNEST_LINE5_GTIN_LABEL, align="L")
            y6 = y5 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            tab = _ARNEST_TAB_MM
            mgt = _ARNEST_LINE6_MGT_DATE_LABEL
            exp = _ARNEST_LINE6_EXPIRY_DATE_LABEL
            cwk = _ARNEST_LINE6_CROSS_WEIGHT_LABEL
            pdf.set_font("PLCalibri", "", fs)
            w_mgt = pdf.get_string_width(mgt)
            w_exp = pdf.get_string_width(exp)
            pdf.set_xy(x0, y6)
            pdf.cell(w_mgt, h_txt, mgt)
            x_exp = x0 + w_mgt + tab + pdf.get_string_width("  ")
            pdf.set_xy(x_exp, y6)
            pdf.cell(w_exp, h_txt, exp)
            x_cwk = x_exp + w_exp + 3 * tab
            pdf.set_xy(x_cwk, y6)
            rem_cwk = max(0.0, max_r - x_cwk)
            if rem_cwk > 0:
                cwk_draw = _arnest_clip_text_to_width_mm(pdf, cwk, rem_cwk)
                pdf.cell(rem_cwk, h_txt, cwk_draw or cwk)
            y7 = y6 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            mfg_raw = _arnest_mfg_date_raw(row)
            exp_raw = _arnest_exp_date_raw(row)
            mfg_txt = _arnest_six_digits_display_ddmmyy(mfg_raw) or "—"
            exp_txt = _arnest_six_digits_display_ddmmyy(exp_raw) or "—"
            w_kg_str = _arnest_format_weight_kg_ru(row.get("pallet_weight_kg", 0))
            pdf.set_font("PLCalibri", "B", fs)
            # Выравниваем строку 7 по тем же колонкам, что и строку 6.
            x_e7 = x_exp
            x_w7 = x_cwk
            rem_m7 = max(0.0, x_e7 - x0)
            rem_e7 = max(0.0, x_w7 - x_e7)
            rem_w7 = max(0.0, max_r - x_w7)
            pdf.set_xy(x0, y7)
            mfg_draw = _arnest_clip_text_to_width_mm(pdf, mfg_txt, rem_m7) if rem_m7 > 0 else ""
            pdf.cell(rem_m7, h_txt, mfg_draw or mfg_txt)
            pdf.set_xy(x_e7, y7)
            exp_draw = _arnest_clip_text_to_width_mm(pdf, exp_txt, rem_e7) if rem_e7 > 0 else ""
            pdf.cell(rem_e7, h_txt, exp_draw or exp_txt)
            pdf.set_xy(x_w7, y7)
            if rem_w7 > 0:
                w_kg_draw = _arnest_clip_text_to_width_mm(pdf, w_kg_str, rem_w7)
                pdf.cell(rem_w7, h_txt, w_kg_draw or w_kg_str)
            y8 = y7 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            boxes_qty = _arnest_format_boxes_qty_for_barcode(
                row.get("pallet_boxes_qty", row.get("pallet_qty", 0))
            )
            try:
                png_boxes = render_code128_barcode_png(boxes_qty, write_text=False)
            except Exception as exc:
                return (
                    None,
                    "barcode_fetch",
                    f"Не удалось сформировать штрих-код количества коробок для паллеты {idx}: {exc}",
                )
            dims_boxes = _barcode_raster_pixel_size(png_boxes)
            if not dims_boxes:
                return None, "pdf_build", ""
            pw2, ph2 = dims_boxes
            h_boxes_mm = (
                _ARNEST_LINE8_BARCODE_W_MM * (ph2 / pw2) * _ARNEST_BARCODE_HEIGHT_SCALE
            )
            pdf.image(
                io.BytesIO(png_boxes),
                x=Align.C,
                y=y8,
                w=_ARNEST_LINE8_BARCODE_W_MM,
                h=h_boxes_mm,
                keep_aspect_ratio=False,
            )
            cap8_h = _arnest_draw_code128_caption_below(
                pdf,
                _ARNEST_SIDE_MARGIN_MM,
                barcode_w_mm,
                y8 + h_boxes_mm,
                boxes_qty,
                body_caption_style=True,
            )
            y9 = y8 + h_boxes_mm + cap8_h + _ARNEST_TEXT_LINE_GAP_MM
            pdf.set_font("PLCalibri", "", fs)
            pdf.set_xy(x0, y9)
            pdf.cell(content_w, h_txt, _ARNEST_LINE9_PALLET_QTY_LABEL, align="L")
            pdf.set_font("PLCalibri", "", fs)
            x_units = x_w7
            pdf.set_xy(x_units, y9)
            pdf.cell(max(0.0, max_r - x_units), h_txt, _ARNEST_LINE9_UNITS_PC_LABEL)
            x_tihi = x_units + pdf.get_string_width(_ARNEST_LINE9_UNITS_PC_LABEL) + tab
            pdf.set_xy(x_tihi, y9)
            pdf.cell(max(0.0, max_r - x_tihi), h_txt, _ARNEST_LINE9_TI_HI_LABEL)
            y10 = y9 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            boxes_str = _arnest_format_boxes_qty_for_barcode(row.get("pallet_boxes_qty", 0))
            rl = int(row.get("row_layout", 0) or 0)
            try:
                boxes_f = float(row.get("pallet_boxes_qty", 0) or 0)
            except (TypeError, ValueError):
                boxes_f = 0.0
            if rl > 0:
                full_rows = int(math.floor(boxes_f / rl + 1e-9))
                ti_hi_str = f"{rl}*{full_rows}"
            else:
                ti_hi_str = "—"
            kor = _ARNEST_LINE10_KOR_LABEL
            pdf.set_font("PLCalibri", "B", fs)
            rem_l10 = max(0.0, x_units - x0 - 1.0)
            pdf.set_xy(x0, y10)
            if rem_l10 > 0:
                boxes_draw = _arnest_clip_text_to_width_mm(pdf, boxes_str, rem_l10)
                pdf.cell(rem_l10, h_txt, boxes_draw or boxes_str)
            else:
                pdf.cell(pdf.get_string_width(boxes_str), h_txt, boxes_str)
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x_units, y10)
            pdf.cell(pdf.get_string_width(kor), h_txt, kor)
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x_tihi, y10)
            rem_t10 = max(0.0, max_r - x_tihi)
            if rem_t10 > 0:
                ti_draw = _arnest_clip_text_to_width_mm(pdf, ti_hi_str, rem_t10)
                pdf.cell(rem_t10, h_txt, ti_draw or ti_hi_str)
            else:
                pdf.cell(pdf.get_string_width(ti_hi_str), h_txt, ti_hi_str)
            y11 = y10 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            rem_star = _arnest_remainder_boxes_star_one(boxes_f, rl)
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x_units, y11)
            pdf.cell(pdf.get_string_width(kor), h_txt, kor)
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x_tihi, y11)
            rem_t11 = max(0.0, max_r - x_tihi)
            if rem_t11 > 0:
                r11 = _arnest_clip_text_to_width_mm(pdf, rem_star, rem_t11)
                pdf.cell(rem_t11, h_txt, r11 or rem_star)
            else:
                pdf.cell(pdf.get_string_width(rem_star), h_txt, rem_star)
            y12 = y11 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            pdf.set_font("PLCalibri", "", fs)
            pdf.set_xy(x0, y12)
            pdf.cell(content_w, h_txt, _ARNEST_LINE12_TOTAL_QUANTITY_LABEL, align="L")
            y13 = y12 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            units_str = _arnest_format_boxes_qty_for_barcode(
                row.get("units_pc", row.get("unitsPc", 0))
            )
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x0, y13)
            rem_13 = content_w
            if rem_13 > 0:
                u13 = _arnest_clip_text_to_width_mm(pdf, units_str, rem_13)
                pdf.cell(rem_13, h_txt, u13 or units_str)
            else:
                pdf.cell(pdf.get_string_width(units_str), h_txt, units_str)
            y14 = y13 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            pallet_bc = build_pallet_barcode_data(
                str(row.get("pallet_number", row.get("palletNumber", "")))
            )
            y_after_line14 = y14
            if pallet_bc:
                try:
                    png_pn = render_code128_barcode_png(pallet_bc, write_text=False)
                except Exception as exc:
                    return (
                        None,
                        "barcode_fetch",
                        f"Не удалось сформировать штрих-код номера паллеты {idx}: {exc}",
                    )
                # Как у штрих-кода коробок (стр. 8): короткий Code128 паллеты иначе получался ниже.
                h_pallet_mm = h_boxes_mm
                pdf.image(
                    io.BytesIO(png_pn),
                    x=x0,
                    y=y14,
                    w=_ARNEST_LINE8_BARCODE_W_MM,
                    h=h_pallet_mm,
                    keep_aspect_ratio=False,
                )
                cap14_h = _arnest_draw_code128_caption_below(
                    pdf,
                    x0,
                    _ARNEST_LINE8_BARCODE_W_MM,
                    y14 + h_pallet_mm,
                    pallet_bc,
                    body_caption_style=True,
                )
                y_after_line14 = y14 + h_pallet_mm + cap14_h + _ARNEST_TEXT_LINE_GAP_MM
            y15 = y_after_line14
            rem_bn = max(0.0, max_r - x_units)
            pdf.set_font("PLCalibri", "", fs)
            pdf.set_xy(x_units, y15)
            pdf.cell(rem_bn, h_txt, _ARNEST_LINE15_BATCH_NUMBER_LABEL, align="L")
            y16 = y15 + h_txt + _ARNEST_TEXT_LINE_GAP_MM
            batch_disp = str(
                row.get("batch_number", row.get("batchNumber", ""))
            ).strip() or "—"
            pdf.set_font("PLCalibri", "B", fs)
            pdf.set_xy(x_units, y16)
            if rem_bn > 0:
                b16 = _arnest_clip_text_to_width_mm(pdf, batch_disp, rem_bn)
                pdf.cell(rem_bn, h_txt, b16 or batch_disp)
            else:
                pdf.cell(pdf.get_string_width(batch_disp), h_txt, batch_disp)
        raw = pdf.output(dest="S")
    except Exception:
        return None, "pdf_build", ""
    if isinstance(raw, (bytes, bytearray)):
        return bytes(raw), "", ""
    if isinstance(raw, str):
        return raw.encode("latin-1"), "", ""
    return None, "pdf_build", ""


def build_blank_arnest_unirus_pallet_sheets_pdf(page_count: int) -> tuple[bytes | None, str]:
    """Один PDF: page_count пустых страниц A4 (пока без содержимого)."""
    if not HAVE_FPDF or FPDF is None:
        return None, "no_fpdf"
    if page_count < 1 or page_count > MAX_PALLET_SHEET_PDF_PAGES:
        return None, "validation"
    try:
        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.set_auto_page_break(False)
        for _ in range(page_count):
            pdf.add_page()
        raw = pdf.output(dest="S")
    except Exception:
        return None, "pdf_build"
    if isinstance(raw, (bytes, bytearray)):
        return bytes(raw), ""
    if isinstance(raw, str):
        return raw.encode("latin-1"), ""
    return None, "pdf_build"


def build_arnest_unirus_pallet_sheets_pdf_bytes(body: dict) -> tuple[bytes | None, str | None, str | None]:
    """Собрать PDF в памяти (тот же ввод, что у POST /api/arnest-unirus-pallet-sheets-pdf).
    Возвращает (pdf_bytes, err_code, err_detail); err_code None при успехе."""
    pallets_raw = body.get("pallets")
    if isinstance(pallets_raw, list) and len(pallets_raw) > 0:
        blob, err, err_detail = build_arnest_unirus_pallet_sheets_pdf_with_barcodes(pallets_raw)
        if err:
            return None, err, err_detail or _arnest_pallet_pdf_error_message(err)
        return blob, None, None
    if isinstance(pallets_raw, list) and len(pallets_raw) == 0:
        return None, "validation_pallets", _arnest_pallet_pdf_error_message("validation_pallets")
    raw = body.get("page_count", body.get("pages", 0))
    try:
        page_count = int(raw)
    except (TypeError, ValueError):
        page_count = 0
    blob, err = build_blank_arnest_unirus_pallet_sheets_pdf(page_count)
    if err:
        return None, err, _arnest_pallet_pdf_error_message(err)
    return blob, None, None


def build_arnest_unirus_pallet_sheets_xlsx_bytes(
    body: dict,
) -> tuple[bytes | None, str | None, str | None]:
    """Сводный XLSX только в памяти (на диск не пишется).

    JSON: { \"pallets\": [...] } — как для PDF; опционально ship_date (строка), order_id (число) — для шапки листа.
    """
    if not HAVE_OPENPYXL or Workbook is None or get_column_letter is None:
        return None, "no_openpyxl", _arnest_pallet_pdf_error_message("no_openpyxl")
    pallets_raw = body.get("pallets")
    ship_date_raw = str(body.get("ship_date") or "").strip()
    order_id_raw = body.get("order_id")
    try:
        order_id_int = int(order_id_raw) if order_id_raw is not None and str(order_id_raw).strip() != "" else None
    except (TypeError, ValueError):
        order_id_int = None
    if not isinstance(pallets_raw, list) or len(pallets_raw) == 0:
        return None, "validation_pallets", _arnest_pallet_pdf_error_message("validation_pallets")
    if len(pallets_raw) > MAX_PALLET_SHEET_PDF_PAGES:
        return None, "validation_pallets", _arnest_pallet_pdf_error_message("validation_pallets")

    validated: list[dict] = []
    for idx, row in enumerate(pallets_raw, start=1):
        if not isinstance(row, dict):
            return None, "validation_pallet", f"Паллета {idx}: ожидался объект с полями."
        _, err_detail = _arnest_line_barcode_data(row)
        if err_detail:
            return None, "validation_pallet", f"Паллета {idx}: {err_detail}."
        validated.append(row)

    try:
        wb = Workbook()
        ws = wb.active
        if ws is None:
            return None, "xlsx_build", _arnest_pallet_pdf_error_message("xlsx_build")
        ws.title = "Паллеты"

        thin = Side(style="thin", color="5A6578")
        grid_border = Border(left=thin, right=thin, top=thin, bottom=thin)
        hdr_fill = PatternFill("solid", fgColor="147A6E")
        hdr_font = Font(bold=True, color="FFFFFF", size=11)
        title_font = Font(bold=True, size=14, color="153045")
        zebra_fill = PatternFill("solid", fgColor="F0F4F8")

        title_parts = ["Арнест Юнирусь — свод по паллетам"]
        if ship_date_raw:
            title_parts.append(f"отгрузка: {ship_date_raw}")
        if order_id_int is not None:
            title_parts.append(f"заказ №{order_id_int}")
        title_text = " · ".join(title_parts)

        last_col = 9
        last_letter = get_column_letter(last_col)
        ws.merge_cells(f"A1:{last_letter}1")
        tcell = ws["A1"]
        tcell.value = title_text
        tcell.font = title_font
        tcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 28

        headers = (
            "Номер паллеты",
            "Артикул",
            "Наименование",
            "Партия",
            "Дата изготовления",
            "Срок годности",
            "Вес брутто, кг",
            "Коробок",
            "Штук",
        )
        hdr_row = 3
        for col_idx, text in enumerate(headers, start=1):
            c = ws.cell(row=hdr_row, column=col_idx, value=text)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = grid_border

        widths = (14, 18, 40, 16, 14, 14, 14, 11, 11)
        for i, wch in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = wch

        data_start = 4
        total_weight = 0.0
        for i, row in enumerate(validated, start=1):
            r = data_start + i - 1
            mfg_raw = _arnest_mfg_date_raw(row)
            exp_raw = _arnest_exp_date_raw(row)
            mfg_disp = _arnest_six_digits_display_ddmmyy(mfg_raw) or "—"
            exp_disp = _arnest_six_digits_display_ddmmyy(exp_raw) or "—"
            try:
                w_kg = float(row.get("pallet_weight_kg", 0) or 0)
            except (TypeError, ValueError):
                w_kg = 0.0
            if not math.isfinite(w_kg) or w_kg < 0:
                w_kg = 0.0
            total_weight += w_kg
            try:
                boxes_f = float(row.get("pallet_boxes_qty", row.get("pallet_qty", 0)) or 0)
            except (TypeError, ValueError):
                boxes_f = 0.0
            try:
                units_f = float(row.get("units_pc", row.get("unitsPc", 0)) or 0)
            except (TypeError, ValueError):
                units_f = 0.0
            boxes_disp = _arnest_format_boxes_qty_for_barcode(boxes_f)
            units_disp = _arnest_format_boxes_qty_for_barcode(units_f)
            pallet_no = str(row.get("pallet_number", row.get("palletNumber", "")) or "").strip()
            article = str(row.get("article", "") or "").strip()
            name = str(row.get("name", row.get("product_name", "")) or "").strip()
            batch = str(row.get("batch_number", row.get("batchNumber", "")) or "").strip()

            vals = (
                pallet_no,
                article,
                name,
                batch,
                mfg_disp,
                exp_disp,
                w_kg,
                boxes_disp,
                units_disp,
            )
            for col_idx, val in enumerate(vals, start=1):
                cell = ws.cell(row=r, column=col_idx, value=val)
                cell.border = grid_border
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                if i % 2 == 0:
                    cell.fill = zebra_fill
            ws.cell(row=r, column=7).number_format = "0.000"

        last_data = data_start + len(validated) - 1
        sum_row = last_data + 1
        ws.merge_cells(start_row=sum_row, start_column=1, end_row=sum_row, end_column=6)
        sum_label = ws.cell(row=sum_row, column=1, value="Итого вес брутто, кг:")
        sum_label.font = Font(bold=True)
        sum_label.alignment = Alignment(horizontal="right", vertical="center")
        sum_label.border = grid_border
        c_w = ws.cell(row=sum_row, column=7, value=round(total_weight, 3))
        c_w.font = Font(bold=True)
        c_w.number_format = "0.000"
        c_w.border = grid_border
        ws.cell(row=sum_row, column=8, value="").border = grid_border
        ws.cell(row=sum_row, column=9, value="").border = grid_border

        ws.freeze_panes = ws.cell(row=data_start, column=1)
        ws.auto_filter.ref = f"A{hdr_row}:{last_letter}{last_data}"

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue(), None, None
    except Exception as exc:
        return None, "xlsx_build", f"{_arnest_pallet_pdf_error_message('xlsx_build')} ({exc})"


def _wrap_pdf_pjl_a4_tray1(pdf: bytes) -> bytes:
    """PJL перед/после PDF для JetDirect: A4, лоток 1, язык PDF."""
    uel = b"\x1b%-12345X"
    crlf = b"\r\n"
    pjl = (
        b'@PJL JOB NAME = "Paletlist pallets"'
        + crlf
        + b"@PJL SET PAPER = A4"
        + crlf
        + b"@PJL SET MEDIASOURCE = TRAY1"
        + crlf
        + b"@PJL ENTER LANGUAGE = PDF"
        + crlf
    )
    end = crlf + b"@PJL EOJ" + crlf + uel
    return uel + pjl + pdf + end


def _get_raw_printer_host_port() -> tuple[str, int | None, str | None]:
    """Хост и порт JetDirect; при ошибке порта — (host, None, текст ошибки)."""
    host = (os.environ.get("PALLET_RAW_PRINT_HOST") or "192.168.1.196").strip()
    port_raw = (os.environ.get("PALLET_RAW_PRINT_PORT") or "9100").strip()
    try:
        return host, int(port_raw), None
    except ValueError:
        return host, None, "Некорректный PALLET_RAW_PRINT_PORT"


def check_raw_pallet_printer_reachable() -> tuple[bool, str, str]:
    """Проверка TCP до принтера без отправки данных на печать.
    Возвращает (успех, \"host:port\", сообщение для пользователя)."""
    host, port, cfg_err = _get_raw_printer_host_port()
    addr = f"{host}:{port}" if port is not None else f"{host}:?"
    if cfg_err:
        return False, addr, cfg_err
    ping_timeout = float((os.environ.get("PALLET_RAW_PRINT_PING_TIMEOUT_SEC") or "5").strip() or "5")
    try:
        with socket.create_connection((host, port), timeout=ping_timeout):
            pass
    except OSError as exc:
        return False, addr, str(exc)
    return True, addr, "TCP-соединение установлено (данные на печать не отправлялись)."


def send_pdf_to_raw_lan_printer(pdf: bytes) -> tuple[bool, str, int]:
    """Отправка PDF на принтер по TCP:9100 (сервер должен видеть IP принтера в LAN).
    Обёртка PJL A4/лоток1 — по умолчанию включена (PALLET_RAW_PRINT_PJL=0 — только сырые байты PDF)."""
    host, port, cfg_err = _get_raw_printer_host_port()
    if cfg_err or port is None:
        return False, cfg_err or "Ошибка настройки порта принтера", 0
    timeout = float((os.environ.get("PALLET_RAW_PRINT_TIMEOUT_SEC") or "25").strip() or "25")
    use_pjl = (os.environ.get("PALLET_RAW_PRINT_PJL", "1") or "1").strip().lower() not in (
        "0",
        "false",
        "no",
        "off",
    )
    payload = _wrap_pdf_pjl_a4_tray1(pdf) if use_pjl else pdf
    try:
        with socket.create_connection((host, port), timeout=timeout) as sock:
            sock.sendall(payload)
    except OSError as exc:
        return False, f"{host}:{port}: {exc}", 0
    return True, f"{host}:{port}", len(payload)


def _parse_orders_detail_id(path: str):
    """Для /api/orders/12 возвращает 12; для списка или чужих путей — None."""
    p = _norm_api_path(path)
    prefix = "/api/orders/"
    if not p.startswith(prefix):
        return None
    tail = p[len(prefix) :]
    if not tail or "/" in tail:
        return None
    try:
        return int(tail)
    except ValueError:
        return None


def _parse_orders_packing_sheets_html_id(path: str) -> int | None:
    """Для GET /api/orders/12/packing-sheets-html возвращает 12."""
    p = _norm_api_path(path)
    prefix = "/api/orders/"
    suffix = "/packing-sheets-html"
    if not p.startswith(prefix) or not p.endswith(suffix):
        return None
    mid = p[len(prefix) : -len(suffix)]
    if not mid or "/" in mid:
        return None
    try:
        return int(mid)
    except ValueError:
        return None


def _parse_orders_packing_sheets_pdf_id(path: str) -> int | None:
    """Для GET /api/orders/12/packing-sheets-pdf возвращает 12."""
    p = _norm_api_path(path)
    prefix = "/api/orders/"
    suffix = "/packing-sheets-pdf"
    if not p.startswith(prefix) or not p.endswith(suffix):
        return None
    mid = p[len(prefix) : -len(suffix)]
    if not mid or "/" in mid:
        return None
    try:
        return int(mid)
    except ValueError:
        return None


def get_connection():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    return con


def _table_columns(cur, table: str):
    return {row[1] for row in cur.execute(f"PRAGMA table_info({table})").fetchall()}


def _migrate_products(cur):
    cols = _table_columns(cur, "products")
    if "pieces_in_box" not in cols:
        cur.execute(
            "ALTER TABLE products ADD COLUMN pieces_in_box INTEGER NOT NULL DEFAULT 0"
        )
    if "row_layout" not in cols:
        cur.execute(
            "ALTER TABLE products ADD COLUMN row_layout INTEGER NOT NULL DEFAULT 0"
        )
    if "max_rows" not in cols:
        cur.execute("ALTER TABLE products ADD COLUMN max_rows INTEGER NOT NULL DEFAULT 0")
    if "box_weight" not in cols:
        cur.execute(
            "ALTER TABLE products ADD COLUMN box_weight REAL NOT NULL DEFAULT 0"
        )
    if "pieces_per_set" not in cols:
        cur.execute(
            "ALTER TABLE products ADD COLUMN pieces_per_set INTEGER NOT NULL DEFAULT 1"
        )
    if "sets_in_box" not in cols:
        cur.execute(
            "ALTER TABLE products ADD COLUMN sets_in_box INTEGER NOT NULL DEFAULT 1"
        )
        for row in cur.execute(
            "SELECT id, pieces_in_box, pieces_per_set FROM products"
        ).fetchall():
            pid = int(row["id"])
            pib = int(row["pieces_in_box"] or 0)
            pps = max(1, int(row["pieces_per_set"] or 1))
            if pps <= 1 or pib <= 0:
                sib = 1
            else:
                sib = pib // pps
                if sib < 1:
                    sib = 1
                if pps > 0 and pib % pps != 0:
                    sib = max(1, int(round(float(pib) / pps)))
            cur.execute(
                "UPDATE products SET sets_in_box = ? WHERE id = ?",
                (sib, pid),
            )
        for row in cur.execute(
            "SELECT id, pieces_in_box, sets_in_box FROM products"
        ).fetchall():
            pid = int(row["id"])
            pib = int(row["pieces_in_box"] or 0)
            sib = max(1, int(row["sets_in_box"] or 1))
            if sib <= 1 or pib <= 0:
                pps = 1
            elif pib % sib == 0:
                pps = max(1, pib // sib)
            else:
                pps = max(1, int(round(float(pib) / sib)))
            cur.execute(
                "UPDATE products SET pieces_per_set = ? WHERE id = ?",
                (pps, pid),
            )


# Хвосты наименований: «, шт» / «, набор» и варианты (запятая ASCII/полношир./„ , опц. точка после шт).
_NAME_SUFFIX_TRAIL_PATTERNS = (
    re.compile(r"\s*[\u002C\uFF0C\u201A]\s*набор\.?\s*$", re.IGNORECASE),
    re.compile(r"\s*[\u002C\uFF0C\u201A]\s*шт\.?\s*$", re.IGNORECASE),
)


def _strip_trailing_name_suffix(name: str) -> str:
    """Убирает хвост «, шт» / «, набор» (и типичные варианты написания) в конце наименования."""
    s = unicodedata.normalize("NFKC", name or "").rstrip()
    changed = True
    while changed and s:
        changed = False
        for pat in _NAME_SUFFIX_TRAIL_PATTERNS:
            nxt = pat.sub("", s).rstrip()
            if nxt != s:
                s = nxt
                changed = True
                break
    return s


def _migrate_strip_trailing_name_suffixes(cur):
    """Идемпотентно чистит products.name и order_items.name от суффиксов «, шт» / «, набор»."""
    for row in cur.execute("SELECT id, name FROM products"):
        old = row["name"] if row["name"] is not None else ""
        new = _strip_trailing_name_suffix(old)
        if new != old:
            cur.execute(
                "UPDATE products SET name = ? WHERE id = ?",
                (new, int(row["id"])),
            )
    if not cur.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='order_items' LIMIT 1"
    ).fetchone():
        return
    for row in cur.execute("SELECT id, name FROM order_items"):
        old = row["name"] if row["name"] is not None else ""
        new = _strip_trailing_name_suffix(old)
        if new != old:
            cur.execute(
                "UPDATE order_items SET name = ? WHERE id = ?",
                (new, int(row["id"])),
            )


def _normalize_packaging(pieces_in_box: int, sets_in_box: int):
    """Возвращает (pib, sib, pps) или dict с error/message при ошибке валидации."""
    pib = max(0, int(pieces_in_box))
    sib = max(1, int(sets_in_box))
    if sib <= 1:
        return (pib, sib, 1)
    if pib <= 0:
        return (pib, sib, 1)
    if pib % sib != 0:
        return {
            "error": "validation",
            "message": "Штук в коробке должно делиться на количество наборов без остатка.",
        }
    pps = max(1, pib // sib)
    return (pib, sib, pps)


def _computed_pieces_per_set_for_api(pib: int, sib: int):
    sib = max(1, int(sib))
    pib = int(pib)
    if sib <= 1 or pib <= 0:
        return 1
    if pib % sib == 0:
        return max(1, pib // sib)
    return round(pib / sib, 2)


def init_db():
    with DB_LOCK:
        con = get_connection()
        cur = con.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS products (
              id INTEGER PRIMARY KEY,
              article TEXT NOT NULL DEFAULT '',
              name TEXT NOT NULL DEFAULT '',
              pieces_in_box INTEGER NOT NULL DEFAULT 0,
              sets_in_box INTEGER NOT NULL DEFAULT 1,
              pieces_per_set INTEGER NOT NULL DEFAULT 1,
              row_layout INTEGER NOT NULL DEFAULT 0,
              max_rows INTEGER NOT NULL DEFAULT 0,
              box_weight REAL NOT NULL DEFAULT 0
            )
            """
        )
        _migrate_products(cur)
        count = cur.execute("SELECT COUNT(*) FROM products").fetchone()[0]
        if count == 0 and JSON_SEED_PATH.exists():
            seed_items = json.loads(JSON_SEED_PATH.read_text(encoding="utf-8"))
            rows = []
            for item in seed_items:
                try:
                    pib = int(item.get("pieces_in_box") or 0)
                    sib = max(1, int(item.get("sets_in_box") or 1))
                    norm = _normalize_packaging(pib, sib)
                    if isinstance(norm, dict):
                        pib, sib, pps = pib, 1, 1
                    else:
                        pib, sib, pps = norm
                    rows.append(
                        (
                            int(item.get("id")),
                            (item.get("article") or "").strip(),
                            _strip_trailing_name_suffix((item.get("name") or "").strip()),
                            pib,
                            sib,
                            pps,
                            int(item.get("row_layout") or 0),
                            int(item.get("max_rows") or 0),
                            float(item.get("box_weight") or 0),
                        )
                    )
                except (TypeError, ValueError):
                    continue
            if rows:
                cur.executemany(
                    """
                    INSERT INTO products (
                      id, article, name,
                      pieces_in_box, sets_in_box, pieces_per_set, row_layout, max_rows, box_weight
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    rows,
                )
        _init_orders_table(cur)
        _migrate_strip_trailing_name_suffixes(cur)
        con.commit()
        con.close()


def _init_orders_table(cur):
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS orders (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          ship_date TEXT NOT NULL DEFAULT '',
          client TEXT NOT NULL DEFAULT '',
          assembled_percent INTEGER NOT NULL DEFAULT 0,
          names TEXT NOT NULL DEFAULT '',
          extra_info TEXT NOT NULL DEFAULT ''
        )
        """
    )
    n = cur.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
    if n == 0:
        samples = [
            (
                "15.05.2026",
                "ООО «Розница Юг»",
                72,
                "Салфетки САЛДЕЗ; Aqua Joy гель Mango 1 л",
                "Паллеты в зоне Б, проверить маркировку",
            ),
            (
                "18.05.2026",
                "ИП Ким А.С.",
                35,
                "Спрей Шаума Kids; Bambolina масло 250 мл",
                "Частичная отгрузка, звонок за 2 ч",
            ),
            (
                "22.05.2026",
                "Сеть «Косметика Плюс»",
                0,
                "Adaly двухфазное средство 120 мл (×24 короба)",
                "Новый заказ, сбор не начат",
            ),
        ]
        cur.executemany(
            """
            INSERT INTO orders (
              ship_date, client, assembled_percent, names, extra_info
            ) VALUES (?, ?, ?, ?, ?)
            """,
            samples,
        )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS order_items (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          order_id INTEGER NOT NULL,
          product_id INTEGER,
          article TEXT NOT NULL DEFAULT '',
          name TEXT NOT NULL DEFAULT '',
          quantity REAL NOT NULL DEFAULT 0,
          unit TEXT NOT NULL DEFAULT 'piece',
          FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE
        )
        """
    )
    order_cols = {r[1] for r in cur.execute("PRAGMA table_info(orders)").fetchall()}
    if "assemble_state" not in order_cols:
        cur.execute(
            "ALTER TABLE orders ADD COLUMN assemble_state TEXT NOT NULL DEFAULT ''"
        )


_MAX_ASSEMBLE_STATE_JSON_BYTES = 2 * 1024 * 1024


def _assemble_state_cell_to_api(value) -> dict | None:
    """JSON из БД → объект для API или None."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    try:
        obj = json.loads(s)
    except json.JSONDecodeError:
        return None
    if not isinstance(obj, dict) or not isinstance(obj.get("pallets"), list):
        return None
    return obj


def patch_order_assembly(order_id: int, body: dict) -> dict:
    """Частичное обновление: assembled_percent и/или assemble_state (распределение по паллетам)."""
    if not isinstance(body, dict):
        return {"error": "validation", "message": "Ожидался JSON-объект."}
    has_pct = "assembled_percent" in body
    has_state = "assemble_state" in body
    if not has_pct and not has_state:
        return {
            "error": "validation",
            "message": "Передайте assembled_percent и/или assemble_state.",
        }
    new_pct = None
    if has_pct:
        try:
            new_pct = max(0, min(100, int(body["assembled_percent"])))
        except (TypeError, ValueError):
            return {"error": "validation", "message": "assembled_percent: нужно целое 0…100."}
    new_state_json = None
    if has_state:
        st = body["assemble_state"]
        if st is None:
            new_state_json = ""
        elif not isinstance(st, dict):
            return {"error": "validation", "message": "assemble_state: ожидался объект или null."}
        else:
            try:
                new_state_json = json.dumps(st, ensure_ascii=False)
            except (TypeError, ValueError):
                return {"error": "validation", "message": "assemble_state: не удалось сериализовать в JSON."}
            if len(new_state_json.encode("utf-8")) > _MAX_ASSEMBLE_STATE_JSON_BYTES:
                return {
                    "error": "validation",
                    "message": "assemble_state слишком большой.",
                }
            pallets = st.get("pallets")
            if not isinstance(pallets, list):
                return {
                    "error": "validation",
                    "message": "assemble_state.pallets: ожидался массив.",
                }
    with DB_LOCK:
        con = get_connection()
        cur = con.cursor()
        row = cur.execute(
            """
            SELECT id, assembled_percent, assemble_state
            FROM orders WHERE id = ?
            """,
            (order_id,),
        ).fetchone()
        if not row:
            con.close()
            return {"error": "not_found", "message": "Заказ не найден."}
        apct = new_pct if new_pct is not None else max(
            0, min(100, int(row["assembled_percent"] or 0))
        )
        if new_state_json is not None:
            ajson = new_state_json
        else:
            ajson = row["assemble_state"] or ""
        cur.execute(
            """
            UPDATE orders SET assembled_percent = ?, assemble_state = ?
            WHERE id = ?
            """,
            (apct, ajson, order_id),
        )
        con.commit()
        con.close()
    return {"ok": True, "id": int(order_id)}


def _format_ship_date_storage(value: str) -> str:
    v = (value or "").strip()
    if len(v) == 10 and v[4] == "-" and v[7] == "-":
        return f"{v[8:10]}.{v[5:7]}.{v[0:4]}"
    return v


def _format_order_names_summary(items):
    unit_ru = {"box": "коробок", "set": "наборов", "piece": "шт"}
    parts = []
    for _pid, article, name, qty, unit in items:
        label = (name or article or "—").strip() or "—"
        u = unit_ru.get(unit, "шт")
        q = int(qty) if qty == int(qty) else qty
        parts.append(f"{label} × {q} {u}")
    return "; ".join(parts)


_ORDER_NAMES_LINE_RE = re.compile(
    r"^(.+?)\s*×\s*([\d.,]+)\s*(коробок|наборов|шт)\s*$",
    re.UNICODE,
)


def _synthetic_items_from_order_names(names: str):
    """Если в order_items нет строк, собираем позиции из сводки names (демо/старые заказы)."""
    if not names or not str(names).strip():
        return []
    out = []
    for part in str(names).split(";"):
        seg = part.strip()
        if not seg:
            continue
        m = _ORDER_NAMES_LINE_RE.match(seg)
        if m:
            label = m.group(1).strip()
            try:
                qty = float(m.group(2).replace(",", "."))
            except (TypeError, ValueError):
                qty = 1.0
            unit = {"коробок": "box", "наборов": "set", "шт": "piece"}.get(m.group(3), "piece")
            out.append(
                {
                    "product_id": None,
                    "article": "",
                    "name": label,
                    "quantity": qty,
                    "unit": unit,
                    "pieces_in_box": 0,
                    "sets_in_box": 1,
                    "pieces_per_set": 1,
                    "row_layout": 0,
                    "max_rows": 0,
                    "box_weight": 0.0,
                }
            )
        else:
            out.append(
                {
                    "product_id": None,
                    "article": "",
                    "name": seg,
                    "quantity": 1.0,
                    "unit": "piece",
                    "pieces_in_box": 0,
                    "sets_in_box": 1,
                    "pieces_per_set": 1,
                    "row_layout": 0,
                    "max_rows": 0,
                    "box_weight": 0.0,
                }
            )
    return out


def _try_resolve_product_id(cur, article: str, name: str):
    """Точное совпадение наименования или артикула с номенклатурой."""
    name_n = _normalize_str(name)
    art_n = _normalize_str(article)
    if name_n:
        row = cur.execute(
            "SELECT id FROM products WHERE name = ? LIMIT 1",
            (name_n,),
        ).fetchone()
        if row:
            return int(row["id"])
    if art_n:
        row = cur.execute(
            "SELECT id FROM products WHERE article = ? LIMIT 1",
            (art_n,),
        ).fetchone()
        if row:
            return int(row["id"])
    return None


def _resolve_normalized_product_ids(normalized, cur):
    """Подставляет product_id и канонические article/name из БД, если id не был передан."""
    out = []
    for pid, article, name, qty, unit in normalized:
        if pid is None:
            rid = _try_resolve_product_id(cur, article, name)
            if rid is None:
                return None, {
                    "error": "validation",
                    "message": "Для позиций без привязки к номенклатуре выберите товар в поле «Наименование» из подсказки или приведите название к точному совпадению с номенклатурой.",
                }
            row = cur.execute(
                "SELECT article, name FROM products WHERE id = ?",
                (rid,),
            ).fetchone()
            if not row:
                return None, {
                    "error": "validation",
                    "message": "Товар не найден в номенклатуре.",
                }
            article = _normalize_str(row["article"])
            name = _normalize_str(row["name"])
            pid = rid
        if not name and not article:
            return None, {
                "error": "validation",
                "message": "Для позиции не указано наименование.",
            }
        out.append((pid, article, name, qty, unit))
    return out, None


def _normalize_order_items_body(body: dict):
    """Общая валидация позиций для создания и обновления заказа."""
    ship_raw = body.get("ship_date", "")
    client_raw = body.get("client", "")
    ship = _format_ship_date_storage(ship_raw)
    client_n = _normalize_str(client_raw)
    if not ship or not client_n:
        return {
            "error": "validation",
            "message": "Укажите дату отгрузки и клиента.",
        }
    raw_items = body.get("items")
    if not isinstance(raw_items, list) or not raw_items:
        return {
            "error": "validation",
            "message": "Добавьте хотя бы одну позицию заказа.",
        }
    normalized = []
    for raw in raw_items:
        if not isinstance(raw, dict):
            continue
        try:
            qty = float(raw.get("quantity", 0))
        except (TypeError, ValueError):
            qty = 0
        unit = (raw.get("unit") or "piece").strip().lower()
        if unit not in ("box", "set", "piece"):
            unit = "piece"
        name = _strip_trailing_name_suffix(_normalize_str(raw.get("name", "")))
        article = _normalize_str(raw.get("article", ""))
        pid_raw = raw.get("product_id")
        pid = None
        if pid_raw is not None and str(pid_raw).strip() != "":
            try:
                pid = int(pid_raw)
            except (TypeError, ValueError):
                pid = None
        if qty <= 0:
            return {
                "error": "validation",
                "message": "Укажите количество больше нуля по каждой позиции.",
            }
        if pid is None and not name and not article:
            return {
                "error": "validation",
                "message": "Для каждой позиции укажите наименование или выберите товар из подсказки.",
            }
        if pid is not None and not name and not article:
            return {
                "error": "validation",
                "message": "Для позиции не указано наименование.",
            }
        normalized.append((pid, article, name, qty, unit))
    if not normalized:
        return {
            "error": "validation",
            "message": "Нет корректных позиций в заказе.",
        }
    return {
        "ok": True,
        "ship": ship,
        "client": client_n,
        "normalized": normalized,
    }


def insert_order_with_items(body: dict):
    pack = _normalize_order_items_body(body)
    if pack.get("error"):
        return pack
    ship = pack["ship"]
    client_n = pack["client"]
    normalized = pack["normalized"]
    with DB_LOCK:
        con = get_connection()
        cur = con.cursor()
        normalized, res_err = _resolve_normalized_product_ids(normalized, cur)
        if res_err:
            con.close()
            return res_err
        names_summary = _format_order_names_summary(normalized)
        cur.execute(
            """
            INSERT INTO orders (ship_date, client, assembled_percent, names, extra_info)
            VALUES (?, ?, 0, ?, '')
            """,
            (ship, client_n, names_summary),
        )
        oid = cur.lastrowid
        for pid, article, name, qty, unit in normalized:
            cur.execute(
                """
                INSERT INTO order_items (order_id, product_id, article, name, quantity, unit)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (oid, pid, article, name, qty, unit),
            )
        con.commit()
        con.close()
    return {"ok": True, "id": int(oid)}


def update_order_with_items(order_id: int, body: dict):
    pack = _normalize_order_items_body(body)
    if pack.get("error"):
        return pack
    ship = pack["ship"]
    client_n = pack["client"]
    normalized = pack["normalized"]
    with DB_LOCK:
        con = get_connection()
        cur = con.cursor()
        normalized, res_err = _resolve_normalized_product_ids(normalized, cur)
        if res_err:
            con.close()
            return res_err
        names_summary = _format_order_names_summary(normalized)
        row = cur.execute(
            "SELECT id, assembled_percent, extra_info, assemble_state FROM orders WHERE id = ?",
            (order_id,),
        ).fetchone()
        if not row:
            con.close()
            return {"error": "not_found", "message": "Заказ не найден."}
        apct = max(0, min(100, int(row["assembled_percent"] or 0)))
        xinfo = row["extra_info"] or ""
        xasm = row["assemble_state"] or ""
        cur.execute(
            """
            UPDATE orders SET ship_date = ?, client = ?, names = ?, assembled_percent = ?, extra_info = ?, assemble_state = ?
            WHERE id = ?
            """,
            (ship, client_n, names_summary, apct, xinfo, xasm, order_id),
        )
        cur.execute("DELETE FROM order_items WHERE order_id = ?", (order_id,))
        for pid, article, name, qty, unit in normalized:
            cur.execute(
                """
                INSERT INTO order_items (order_id, product_id, article, name, quantity, unit)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (order_id, pid, article, name, qty, unit),
            )
        con.commit()
        con.close()
    return {"ok": True, "id": int(order_id)}


def _order_item_row_to_dict(ir):
    """Строка JOIN order_items + products → dict как в fetch_order_detail."""
    pid = ir["product_id"]
    return {
        "product_id": int(pid) if pid is not None else None,
        "article": ir["article"] or "",
        "name": ir["name"] or "",
        "quantity": float(ir["quantity"] or 0),
        "unit": (ir["unit"] or "piece").strip().lower(),
        "pieces_in_box": max(0, int(ir["pieces_in_box"] or 0)),
        "sets_in_box": max(1, int(ir["sets_in_box"] or 1)),
        "pieces_per_set": max(1, int(ir["pieces_per_set"] or 1)),
        "row_layout": max(0, int(ir["row_layout"] or 0)),
        "max_rows": max(0, int(ir["max_rows"] or 0)),
        "box_weight": float(ir["box_weight"] or 0),
    }


def fetch_order_detail(order_id: int):
    with DB_LOCK:
        con = get_connection()
        row = con.execute(
            """
            SELECT id, ship_date, client, assembled_percent, names, extra_info, assemble_state
            FROM orders WHERE id = ?
            """,
            (order_id,),
        ).fetchone()
        if not row:
            con.close()
            return None
        item_rows = con.execute(
            """
            SELECT oi.product_id AS product_id, oi.article AS article, oi.name AS name,
                   oi.quantity AS quantity, oi.unit AS unit,
                   p.pieces_in_box AS pieces_in_box,
                   p.sets_in_box AS sets_in_box,
                   p.pieces_per_set AS pieces_per_set,
                   p.row_layout AS row_layout,
                   p.max_rows AS max_rows,
                   p.box_weight AS box_weight
            FROM order_items oi
            LEFT JOIN products p ON p.id = oi.product_id
            WHERE oi.order_id = ?
            ORDER BY oi.id
            """,
            (order_id,),
        ).fetchall()
        con.close()
    items = [_order_item_row_to_dict(ir) for ir in item_rows]
    if not items and (row["names"] or "").strip():
        items = _synthetic_items_from_order_names(row["names"] or "")
    return {
        "id": int(row["id"]),
        "ship_date": row["ship_date"] or "",
        "client": row["client"] or "",
        "assembled_percent": max(0, min(100, int(row["assembled_percent"] or 0))),
        "names": row["names"] or "",
        "extra_info": row["extra_info"] or "",
        "assemble_state": _assemble_state_cell_to_api(row["assemble_state"]),
        "items": items,
    }


def delete_order(order_id: int) -> bool:
    """Удаляет заказ; позиции order_items удаляются каскадом.
    PDF/XLSX паллетных листов на сервере не кэшируются — генерируются в памяти по запросу, отдельных файлов под заказ нет."""
    with DB_LOCK:
        con = get_connection()
        cur = con.cursor()
        cur.execute("DELETE FROM orders WHERE id = ?", (order_id,))
        n = cur.rowcount
        con.commit()
        con.close()
    return n > 0


def fetch_orders():
    with DB_LOCK:
        con = get_connection()
        rows = con.execute(
            """
            SELECT id, ship_date, client, assembled_percent, names, extra_info, assemble_state
            FROM orders ORDER BY id DESC
            """
        ).fetchall()
        ids = [int(r["id"]) for r in rows]
        items_by_oid = {oid: [] for oid in ids}
        if ids:
            placeholders = ",".join("?" * len(ids))
            item_rows = con.execute(
                f"""
                SELECT oi.order_id AS order_id, oi.product_id AS product_id,
                       oi.article AS article, oi.name AS name,
                       oi.quantity AS quantity, oi.unit AS unit,
                       p.pieces_in_box AS pieces_in_box,
                       p.sets_in_box AS sets_in_box,
                       p.pieces_per_set AS pieces_per_set,
                       p.row_layout AS row_layout,
                       p.max_rows AS max_rows,
                       p.box_weight AS box_weight
                FROM order_items oi
                LEFT JOIN products p ON p.id = oi.product_id
                WHERE oi.order_id IN ({placeholders})
                ORDER BY oi.order_id, oi.id
                """,
                ids,
            ).fetchall()
            for ir in item_rows:
                oid = int(ir["order_id"])
                if oid in items_by_oid:
                    items_by_oid[oid].append(_order_item_row_to_dict(ir))
        con.close()
    out = []
    for row in rows:
        oid = int(row["id"])
        items = items_by_oid.get(oid, [])
        if not items and (row["names"] or "").strip():
            items = _synthetic_items_from_order_names(row["names"] or "")
        out.append(
            {
                "id": oid,
                "ship_date": row["ship_date"] or "",
                "client": row["client"] or "",
                "assembled_percent": max(0, min(100, int(row["assembled_percent"] or 0))),
                "names": row["names"] or "",
                "extra_info": row["extra_info"] or "",
                "assemble_state": _assemble_state_cell_to_api(row["assemble_state"]),
                "items": items,
            }
        )
    return out


def fetch_nomenclature():
    """Список номенклатуры; имена без хвоста «, шт» / «, набор», при расхождении — правка в БД."""
    with DB_LOCK:
        con = get_connection()
        cur = con.cursor()
        rows = cur.execute(
            """
            SELECT id, article, name,
                   pieces_in_box, sets_in_box, pieces_per_set, row_layout, max_rows, box_weight
            FROM products ORDER BY id
            """
        ).fetchall()
        updates: list[tuple[str, int]] = []
        out = []
        for row in rows:
            pib = int(row["pieces_in_box"] or 0)
            sib = max(1, int(row["sets_in_box"] or 1))
            pps = _computed_pieces_per_set_for_api(pib, sib)
            raw_name = row["name"] if row["name"] is not None else ""
            clean_name = _strip_trailing_name_suffix(raw_name)
            if clean_name != raw_name:
                updates.append((clean_name, int(row["id"])))
            out.append(
                {
                    "id": row["id"],
                    "article": row["article"] or "",
                    "name": clean_name,
                    "pieces_in_box": pib,
                    "sets_in_box": sib,
                    "pieces_per_set": pps,
                    "row_layout": int(row["row_layout"] or 0),
                    "max_rows": int(row["max_rows"] or 0),
                    "box_weight": float(row["box_weight"] or 0),
                }
            )
        for clean_name, rid in updates:
            cur.execute(
                "UPDATE products SET name = ? WHERE id = ?",
                (clean_name, rid),
            )
        if updates:
            con.commit()
        con.close()
    return out


def update_nomenclature_row(
    row_id: int,
    article: str,
    name: str,
    pieces_in_box: int = 0,
    sets_in_box: int = 1,
    row_layout: int = 0,
    max_rows: int = 0,
    box_weight: float = 0,
):
    norm = _normalize_packaging(pieces_in_box, sets_in_box)
    if isinstance(norm, dict):
        return norm
    pib, sib, pps = norm
    with DB_LOCK:
        con = get_connection()
        cur = con.cursor()
        cur.execute(
            """
            UPDATE products SET
              article = ?, name = ?,
              pieces_in_box = ?, sets_in_box = ?, pieces_per_set = ?, row_layout = ?, max_rows = ?, box_weight = ?
            WHERE id = ?
            """,
            (
                article.strip(),
                _strip_trailing_name_suffix(name.strip()),
                pib,
                sib,
                pps,
                int(row_layout),
                int(max_rows),
                float(box_weight),
                row_id,
            ),
        )
        affected = cur.rowcount
        con.commit()
        con.close()
    return True if affected > 0 else False


def soft_delete_row(row_id: int):
    return update_nomenclature_row(row_id, "", "", 0, 1, 0, 0, 0) is True


def _normalize_str(value: str) -> str:
    return (value or "").strip()


def find_nomenclature_create_conflicts(article: str, name: str):
    article_n = _normalize_str(article)
    name_n = _normalize_str(name)
    conflicts = []
    if not article_n or not name_n:
        return {
            "error": "validation",
            "message": "Укажите непустой артикул и наименование.",
        }
    art_cf = article_n.casefold()
    name_cf = name_n.casefold()
    for item in fetch_nomenclature():
        ea = _normalize_str(item["article"])
        en = _normalize_str(item["name"])
        if not ea and not en:
            continue
        row_id = item["id"]
        if art_cf == ea.casefold():
            conflicts.append(
                {
                    "kind": "article",
                    "you": article_n,
                    "existing_id": row_id,
                    "existing_article": ea,
                    "existing_name": en or "—",
                }
            )
        if name_cf == en.casefold():
            conflicts.append(
                {
                    "kind": "name",
                    "you": name_n,
                    "existing_id": row_id,
                    "existing_article": ea or "—",
                    "existing_name": en,
                }
            )
    if conflicts:
        return {"error": "duplicate", "conflicts": conflicts}
    return None


def insert_nomenclature_row(
    article: str,
    name: str,
    pieces_in_box: int,
    sets_in_box: int,
    row_layout: int,
    max_rows: int,
    box_weight: float,
):
    article_n = _normalize_str(article)
    name_n = _strip_trailing_name_suffix(_normalize_str(name))
    conflict = find_nomenclature_create_conflicts(article_n, name_n)
    if conflict:
        return conflict
    norm = _normalize_packaging(pieces_in_box, sets_in_box)
    if isinstance(norm, dict):
        return norm
    pib, sib, pps = norm
    with DB_LOCK:
        con = get_connection()
        cur = con.cursor()
        next_id = cur.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM products").fetchone()[0]
        cur.execute(
            """
            INSERT INTO products (
              id, article, name,
              pieces_in_box, sets_in_box, pieces_per_set, row_layout, max_rows, box_weight
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(next_id),
                article_n,
                name_n,
                pib,
                sib,
                pps,
                int(row_layout),
                int(max_rows),
                float(box_weight),
            ),
        )
        con.commit()
        con.close()
    return {"ok": True, "id": int(next_id)}


class ApiHandler(BaseHTTPRequestHandler):
    def _send_json(self, status: int, data):
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _read_json_body(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            return {}
        raw = self.rfile.read(length).decode("utf-8")
        return json.loads(raw) if raw else {}

    def _respond_arnest_unirus_pallet_sheets_pdf(self, body: dict):
        """POST JSON: { \"pallets\": [...] } с полями article, batch_number, manufacturing_date_raw, expiry_date_raw
        (даты — 6 цифр ДДММГГ), опционально pallet_number для штрих-кода строки 14 (1500000+суффикс),
        либо только { \"page_count\": N } — пустые страницы A4."""
        blob, err, err_detail = build_arnest_unirus_pallet_sheets_pdf_bytes(body)
        if err:
            status_map = {
                "validation_pallets": 400,
                "validation_pallet": 400,
                "validation": 400,
                "no_fpdf": 503,
                "no_barcode": 503,
                "no_line2_font": 503,
                "pdf_build": 500,
                "barcode_fetch": 500,
            }
            status = status_map.get(err, 500)
            msg = err_detail or _arnest_pallet_pdf_error_message(err)
            self._send_json(status, {"error": err, "message": msg})
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/pdf")
        self.send_header("Content-Length", str(len(blob)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(blob)

    def _respond_arnest_unirus_pallet_sheets_xlsx(self, body: dict):
        """POST JSON: { \"pallets\": [...] }; опционально ship_date, order_id — для шапки. Файл только в ответе, на диск не пишется."""
        blob, err, err_detail = build_arnest_unirus_pallet_sheets_xlsx_bytes(body)
        if err:
            status_map = {
                "validation_pallets": 400,
                "validation_pallet": 400,
                "no_openpyxl": 503,
                "xlsx_build": 500,
            }
            status = status_map.get(err, 500)
            msg = err_detail or _arnest_pallet_pdf_error_message(err)
            self._send_json(status, {"error": err, "message": msg})
            return
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Length", str(len(blob)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(blob)

    def _respond_print_arnest_unirus_pallet_sheets_raw(self, body: dict):
        """Тот же JSON, что для PDF; сервер отправляет сформированный PDF на принтер (TCP JetDirect)."""
        blob, err, err_detail = build_arnest_unirus_pallet_sheets_pdf_bytes(body)
        if err:
            status_map = {
                "validation_pallets": 400,
                "validation_pallet": 400,
                "validation": 400,
                "no_fpdf": 503,
                "no_barcode": 503,
                "no_line2_font": 503,
                "pdf_build": 500,
                "barcode_fetch": 500,
            }
            status = status_map.get(err, 500)
            msg = err_detail or _arnest_pallet_pdf_error_message(err)
            self._send_json(status, {"error": err, "message": msg})
            return
        ok, info, nbytes = send_pdf_to_raw_lan_printer(blob)
        if not ok:
            self._send_json(
                502,
                {"error": "print_transport", "message": info or "Ошибка отправки на принтер"},
            )
            return
        self._send_json(
            200,
            {"ok": True, "printer": info, "sent_bytes": nbytes},
        )

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        packing_html_oid = _parse_orders_packing_sheets_html_id(path)
        if packing_html_oid is not None:
            detail = fetch_order_detail(packing_html_oid)
            if detail is None:
                self._send_json(404, {"error": "Not found"})
                return
            if build_generic_packing_sheets_html is None:
                self._send_json(
                    503,
                    {"error": "module", "message": "Модуль упаковочных листов недоступен."},
                )
                return
            html, err = build_generic_packing_sheets_html(detail)
            if err:
                known = {"arnest_client", "no_assembly", "no_pallets"}
                if err in known:
                    status_map = {"arnest_client": 400, "no_assembly": 400, "no_pallets": 400}
                    status = status_map.get(err, 400)
                    msg = (
                        generic_packing_error_message(err)
                        if generic_packing_error_message
                        else str(err)
                    )
                    self._send_json(status, {"error": err, "message": msg})
                    return
                self._send_json(500, {"error": "packing_sheet", "message": str(err)})
                return
            b = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(b)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(b)
            return
        packing_oid = _parse_orders_packing_sheets_pdf_id(path)
        if packing_oid is not None:
            detail = fetch_order_detail(packing_oid)
            if detail is None:
                self._send_json(404, {"error": "Not found"})
                return
            if build_generic_packing_sheets_pdf_bytes is None:
                self._send_json(
                    503,
                    {"error": "module", "message": "Модуль упаковочных листов недоступен."},
                )
                return
            pdf_blob, err = build_generic_packing_sheets_pdf_bytes(detail)
            if err:
                known = {"arnest_client", "no_assembly", "no_pallets", "no_pdf_engine", "pdf_empty"}
                if err in known:
                    status_map = {
                        "arnest_client": 400,
                        "no_assembly": 400,
                        "no_pallets": 400,
                        "no_pdf_engine": 503,
                        "pdf_empty": 500,
                    }
                    status = status_map.get(err, 400)
                    msg = (
                        generic_packing_error_message(err)
                        if generic_packing_error_message
                        else str(err)
                    )
                    self._send_json(status, {"error": err, "message": msg})
                    return
                if err.startswith("pdf_render:"):
                    msg = (
                        generic_packing_error_message(err)
                        if generic_packing_error_message
                        else err
                    )
                    self._send_json(500, {"error": "pdf_render", "message": msg})
                    return
                self._send_json(500, {"error": "packing_sheet", "message": str(err)})
                return
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Length", str(len(pdf_blob)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(pdf_blob)
            return
        oid = _parse_orders_detail_id(path)
        if oid is not None:
            detail = fetch_order_detail(oid)
            if detail is None:
                self._send_json(404, {"error": "Not found"})
                return
            self._send_json(200, detail)
            return
        if _is_orders_list_path(path):
            self._send_json(200, fetch_orders())
            return
        if _is_pallet_printer_raw_ping_path(path):
            ok, addr, msg = check_raw_pallet_printer_reachable()
            self._send_json(
                200,
                {"reachable": ok, "printer": addr, "message": msg},
            )
            return
        if not _is_nomenclature_list_path(path):
            self._send_json(404, {"error": "Not found"})
            return
        items = fetch_nomenclature()
        qs = parse_qs(parsed.query)
        if "q" in qs:
            query = (qs.get("q", [""])[0] or "").strip().lower()
            if query:
                items = [
                    item
                    for item in items
                    if query in item["article"].lower() or query in item["name"].lower()
                ]
            else:
                items = []
        self._send_json(200, items)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if _is_print_arnest_unirus_pallet_sheets_raw_path(path):
            try:
                body = self._read_json_body()
            except json.JSONDecodeError:
                self._send_json(400, {"error": "Invalid JSON"})
                return
            self._respond_print_arnest_unirus_pallet_sheets_raw(body)
            return
        if _is_arnest_unirus_pallet_sheets_xlsx_path(path):
            try:
                body = self._read_json_body()
            except json.JSONDecodeError:
                self._send_json(400, {"error": "Invalid JSON"})
                return
            self._respond_arnest_unirus_pallet_sheets_xlsx(body)
            return
        if _is_arnest_unirus_pallet_sheets_pdf_path(path):
            try:
                body = self._read_json_body()
            except json.JSONDecodeError:
                self._send_json(400, {"error": "Invalid JSON"})
                return
            self._respond_arnest_unirus_pallet_sheets_pdf(body)
            return
        if _is_orders_list_path(path):
            try:
                body = self._read_json_body()
            except json.JSONDecodeError:
                self._send_json(400, {"error": "Invalid JSON"})
                return
            try:
                result = insert_order_with_items(body)
            except sqlite3.Error as exc:
                self._send_json(500, {"error": "database", "message": str(exc)})
                return
            err = result.get("error")
            if err == "validation":
                self._send_json(400, result)
                return
            self._send_json(201, result)
            return
        if not _is_nomenclature_list_path(parsed.path):
            self._send_json(
                404,
                {
                    "error": "Not found",
                    "message": "Неизвестный POST-маршрут. Обновите код api_server и выполните sudo systemctl restart paletlist-api.",
                },
            )
            return
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json(400, {"error": "Invalid JSON"})
            return

        def _int_body(key: str) -> int:
            try:
                return int(body.get(key, 0))
            except (TypeError, ValueError):
                return 0

        def _float_body(key: str) -> float:
            try:
                return float(body.get(key, 0))
            except (TypeError, ValueError):
                return 0.0

        def _sets_in_box_body() -> int:
            try:
                v = int(body.get("sets_in_box", 1))
            except (TypeError, ValueError):
                return 1
            return max(1, v)

        try:
            result = insert_nomenclature_row(
                body.get("article", ""),
                body.get("name", ""),
                _int_body("pieces_in_box"),
                _sets_in_box_body(),
                _int_body("row_layout"),
                _int_body("max_rows"),
                _float_body("box_weight"),
            )
        except sqlite3.Error as exc:
            self._send_json(
                500,
                {"error": "database", "message": str(exc)},
            )
            return
        err = result.get("error")
        if err == "validation":
            self._send_json(400, result)
            return
        if err == "duplicate":
            self._send_json(409, result)
            return
        self._send_json(201, result)

    def do_PUT(self):
        parsed = urlparse(self.path)
        path = parsed.path
        oid = _parse_orders_detail_id(path)
        if oid is not None:
            try:
                body = self._read_json_body()
            except json.JSONDecodeError:
                self._send_json(400, {"error": "Invalid JSON"})
                return
            try:
                result = update_order_with_items(oid, body)
            except sqlite3.Error as exc:
                self._send_json(500, {"error": "database", "message": str(exc)})
                return
            err = result.get("error")
            if err == "validation":
                self._send_json(400, result)
                return
            if err == "not_found":
                self._send_json(404, result)
                return
            self._send_json(200, result)
            return
        if not parsed.path.startswith("/api/nomenclature/"):
            self._send_json(404, {"error": "Not found"})
            return
        try:
            row_id = int(parsed.path.rsplit("/", 1)[-1])
        except ValueError:
            self._send_json(400, {"error": "Invalid id"})
            return
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json(400, {"error": "Invalid JSON"})
            return
        def _int_body(key: str) -> int:
            try:
                return int(body.get(key, 0))
            except (TypeError, ValueError):
                return 0

        def _float_body(key: str) -> float:
            try:
                return float(body.get(key, 0))
            except (TypeError, ValueError):
                return 0.0

        def _sets_in_box_body() -> int:
            try:
                v = int(body.get("sets_in_box", 1))
            except (TypeError, ValueError):
                return 1
            return max(1, v)

        ok = update_nomenclature_row(
            row_id,
            body.get("article", ""),
            body.get("name", ""),
            _int_body("pieces_in_box"),
            _sets_in_box_body(),
            _int_body("row_layout"),
            _int_body("max_rows"),
            _float_body("box_weight"),
        )
        if isinstance(ok, dict) and ok.get("error") == "validation":
            self._send_json(400, ok)
            return
        if not ok:
            self._send_json(404, {"error": "Item not found"})
            return
        self._send_json(200, {"ok": True})

    def do_PATCH(self):
        parsed = urlparse(self.path)
        path = parsed.path
        oid = _parse_orders_detail_id(path)
        if oid is None:
            self._send_json(404, {"error": "Not found"})
            return
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json(400, {"error": "Invalid JSON"})
            return
        try:
            result = patch_order_assembly(oid, body)
        except sqlite3.Error as exc:
            self._send_json(500, {"error": "database", "message": str(exc)})
            return
        err = result.get("error")
        if err == "validation":
            self._send_json(400, result)
            return
        if err == "not_found":
            self._send_json(404, result)
            return
        self._send_json(200, result)

    def do_DELETE(self):
        parsed = urlparse(self.path)
        path = parsed.path
        oid = _parse_orders_detail_id(path)
        if oid is not None:
            try:
                ok = delete_order(oid)
            except sqlite3.Error as exc:
                self._send_json(500, {"error": "database", "message": str(exc)})
                return
            if not ok:
                self._send_json(404, {"error": "not_found", "message": "Заказ не найден."})
                return
            self._send_json(200, {"ok": True})
            return
        if not parsed.path.startswith("/api/nomenclature/"):
            self._send_json(404, {"error": "Not found"})
            return
        try:
            row_id = int(parsed.path.rsplit("/", 1)[-1])
        except ValueError:
            self._send_json(400, {"error": "Invalid id"})
            return
        ok = soft_delete_row(row_id)
        if not ok:
            self._send_json(404, {"error": "Item not found"})
            return
        self._send_json(200, {"ok": True})

    def log_message(self, format, *args):
        return


def main():
    server = ThreadingHTTPServer((HOST, PORT), ApiHandler)
    print(f"API server listening on http://{HOST}:{PORT}")
    server.serve_forever()


init_db()

if __name__ == "__main__":
    main()

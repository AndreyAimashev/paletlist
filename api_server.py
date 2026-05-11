#!/usr/bin/env python3
import io
import json
import re
import sqlite3
import threading
import urllib.error
import urllib.request
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse
from xml.sax.saxutils import escape as _xml_escape_text

try:
    from docx import Document as DocxDocument
    from docx.shared import Mm

    HAVE_DOCX = True
except ImportError:
    HAVE_DOCX = False
    DocxDocument = None  # type: ignore
    Mm = None  # type: ignore


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "warehouse.db"
JSON_SEED_PATH = BASE_DIR / "nomenclature.json"
HOST = "127.0.0.1"
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


def _unirus_template_path() -> Path:
    return BASE_DIR / "templates" / "Unirus.docx"


def _is_unirus_template_path(path: str) -> bool:
    return _norm_api_path(path) == "/templates/Unirus.docx"


def _is_unirus_pallet_sheet_path(path: str) -> bool:
    return _norm_api_path(path) == "/api/unirus-pallet-sheet"


UNIRUS_PALLET_BARCODE_PLACEHOLDER = "###ШТРИХКОД_ПАЛЛЕТА###"
UNIRUS_BOX_BARCODE_PLACEHOLDER = "###ШТРИХКОД_КОРОБОК###"
UNIRUS_COMPLEX_BARCODE_PLACEHOLDER = "###СЛОЖНЫЙ_ШТРИХКОД###"
UNIRUS_PH_MFG = "###ДАТА_ИЗГОТОВЛЕНИЯ###"
UNIRUS_PH_EXP = "###ДАТА_ГОДНОСТИ###"
UNIRUS_PH_BOXES_IN_ROWS = "###КОРОБОК_В_РЯДАХ###"
UNIRUS_PH_FULL_ROWS = "###ПОЛНЫХ_РЯДОВ###"
UNIRUS_PH_EXTRA_BOXES = "###КОРОБОК_СВЕРХ###"
UNIRUS_PH_TOTAL_BOXES = "###ВСЕГО_КОРОБОК###"
UNIRUS_PH_TOTAL_PIECES = "###ВСЕГО_ШТУК###"
UNIRUS_PH_PALLET_WEIGHT = "###ВЕС_ПАЛЛЕТА###"
UNIRUS_PH_BATCH = "###НОМЕР_ПАРТИИ###"
UNIRUS_PH_PRODUCT_NAME = "###\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435###"
UNIRUS_OPTIONAL_BODY_KEYS = frozenset(
    {
        "manufacturing_date_raw",
        "expiry_date_raw",
        "boxes_in_full_rows",
        "full_rows_count",
        "extra_boxes",
        "pieces_per_box",
        "box_weight",
        "batch_number",
        "product_name",
    }
)
# В шаблоне «Артикул» часто разбит на три <w:r>: ### | Артикул | ###
UNIRUS_ARTICLE_CORE = "\u0410\u0440\u0442\u0438\u043a\u0443\u043b"
# Без re.DOTALL: не «проглатываем» лишние узлы между run'ами.
_UNIRUS_TRIPLET_ARTICLE_RE = re.compile(
    r"(?P<a1><w:r\b[^>]*>)(?P<apr1><w:rPr>.*?</w:rPr>)<w:t>###</w:t></w:r>"
    r"(?P<a2><w:r\b[^>]*>)(?P<apr2><w:rPr>.*?</w:rPr>)?<w:t>"
    + re.escape(UNIRUS_ARTICLE_CORE)
    + r"</w:t></w:r>"
    r"(?P<a3><w:r\b[^>]*>)(?P<apr3><w:rPr>.*?</w:rPr>)<w:t>###</w:t></w:r>",
)
UNIRUS_BARCODE_PREFIX = "1500000"
BARCODE_FETCH_TIMEOUT_S = 20
BARCODE_USER_AGENT = "paletlist-api/1"


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
    return UNIRUS_BARCODE_PREFIX + _format_pallet_suffix_vba(n)


def fetch_barcode_png(barcode_data: str) -> bytes:
    url = (
        "https://barcode.tec-it.com/barcode.ashx?data="
        + quote(barcode_data, safe="")
        + "&code=Code128&translate-esc=on"
    )
    req = urllib.request.Request(url, headers={"User-Agent": BARCODE_USER_AGENT})
    with urllib.request.urlopen(req, timeout=BARCODE_FETCH_TIMEOUT_S) as resp:
        return resp.read()


def _unirus_strip_str(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _unirus_extra_from_body(body: dict) -> dict:
    out: dict[str, str] = {}
    for k in UNIRUS_OPTIONAL_BODY_KEYS:
        if k not in body:
            continue
        v = body.get(k)
        if v is None:
            continue
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            out[k] = str(v).strip()
        else:
            out[k] = _unirus_strip_str(v)
    return out


def _unirus_val(s: str) -> float:
    t = _unirus_strip_str(s).replace(",", ".")
    if not t:
        return 0.0
    try:
        return float(t)
    except ValueError:
        return 0.0


def format_date_for_sheet_vba(date_input: str) -> str:
    """Как FormatDateForSheet в макросе Word (6 цифр → YY.MM.20XX)."""
    raw = _unirus_strip_str(date_input)
    if not raw:
        return ""
    if len(raw) == 6 and raw.isdigit():
        return f"{raw[:2]}.{raw[2:4]}.20{raw[4:6]}"
    return raw


def convert_date_block_for_complex_vba(date_input: str) -> str:
    """Как ConvertDateToYYMMDD: YYMMDD → DDMMYY (три пары переставлены)."""
    raw = _unirus_strip_str(date_input)
    if not raw or len(raw) != 6 or not raw.isdigit():
        return ""
    return raw[4:6] + raw[2:4] + raw[:2]


def final_batch_vba(batch: str, mfg_raw: str) -> str:
    """Как в ProcessBarcodeData для ###НОМЕР_ПАРТИИ###."""
    b = _unirus_strip_str(batch)
    if not b:
        return ""
    mfg = _unirus_strip_str(mfg_raw)
    if len(b) <= 2 and len(mfg) == 6 and mfg.isdigit():
        try:
            bn = int(float(b.replace(",", ".")))
        except ValueError:
            return b
        return convert_date_block_for_complex_vba(mfg) + f"{bn % 100:02d}"
    return b


def first_number_for_complex_barcode(article: str) -> str:
    """Аналог FindFirstNumber: первое целое число из артикула (цифры подряд)."""
    t = _unirus_strip_str(article)
    if not t:
        return ""
    if t.isdigit():
        return t
    m = re.search(r"\d+", t)
    return m.group(0) if m else ""


def _unirus_cstr_number(n: float) -> str:
    if abs(n - round(n)) < 1e-9:
        return str(int(round(n)))
    return str(n).replace(",", ".")


def _compute_unirus_text_replacements(
    article: str, extra: dict | None
) -> tuple[dict[str, str], str, str]:
    """
    (текстовые замены, строка для сложного Code128, total_boxes_str).
    Пустое значение в словаре = не подставлять (как InsertDateAtPlaceholder в макросе).
    """
    x = extra or {}
    mfg_raw = _unirus_strip_str(x.get("manufacturing_date_raw", ""))
    exp_raw = _unirus_strip_str(x.get("expiry_date_raw", ""))
    boxes_in_row = _unirus_strip_str(x.get("boxes_in_full_rows", ""))
    rows_count = _unirus_strip_str(x.get("full_rows_count", ""))
    extra_boxes = _unirus_strip_str(x.get("extra_boxes", ""))
    pieces = _unirus_strip_str(x.get("pieces_per_box", ""))
    weight = _unirus_strip_str(x.get("box_weight", ""))
    batch_raw = _unirus_strip_str(x.get("batch_number", ""))
    product_name = _unirus_strip_str(x.get("product_name", ""))

    out: dict[str, str] = {}
    if mfg_raw:
        out[UNIRUS_PH_MFG] = format_date_for_sheet_vba(mfg_raw)
    if exp_raw:
        out[UNIRUS_PH_EXP] = format_date_for_sheet_vba(exp_raw)
    if boxes_in_row:
        out[UNIRUS_PH_BOXES_IN_ROWS] = boxes_in_row
    if rows_count:
        out[UNIRUS_PH_FULL_ROWS] = rows_count
    if extra_boxes:
        out[UNIRUS_PH_EXTRA_BOXES] = extra_boxes

    total_boxes_str = ""
    if boxes_in_row and rows_count and extra_boxes:
        tb = _unirus_val(boxes_in_row) * _unirus_val(rows_count) + _unirus_val(extra_boxes)
        total_boxes_str = _unirus_cstr_number(tb)
    if total_boxes_str:
        out[UNIRUS_PH_TOTAL_BOXES] = total_boxes_str

    total_pieces_str = ""
    if total_boxes_str and pieces:
        tp = _unirus_val(total_boxes_str) * _unirus_val(pieces)
        total_pieces_str = _unirus_cstr_number(tp)
    if total_pieces_str:
        out[UNIRUS_PH_TOTAL_PIECES] = total_pieces_str

    total_weight_str = ""
    if total_boxes_str and weight:
        tw = _unirus_val(total_boxes_str) * _unirus_val(weight)
        total_weight_str = _unirus_cstr_number(tw)
    if total_weight_str:
        out[UNIRUS_PH_PALLET_WEIGHT] = total_weight_str

    final_b = final_batch_vba(batch_raw, mfg_raw)
    if final_b:
        out[UNIRUS_PH_BATCH] = final_b

    if product_name:
        out[UNIRUS_PH_PRODUCT_NAME] = product_name

    first_n = first_number_for_complex_barcode(article)
    mfg_blk = convert_date_block_for_complex_vba(mfg_raw)
    exp_blk = convert_date_block_for_complex_vba(exp_raw)
    complex_data = ""
    if first_n and final_b and mfg_blk and exp_blk:
        complex_data = f"{first_n} {final_b} {mfg_blk} {exp_blk}"
    return out, complex_data, total_boxes_str


def _replace_paragraph_placeholder(paragraph, placeholder: str, image_bytes: bytes) -> None:
    before, _, after = paragraph.text.partition(placeholder)
    p_elm = paragraph._element
    for child in list(p_elm):
        p_elm.remove(child)
    if before:
        paragraph.add_run(before)
    run = paragraph.add_run()
    run.add_picture(io.BytesIO(image_bytes), width=Mm(52))
    if after:
        paragraph.add_run(after)


def _walk_unirus_replace_barcode(doc, placeholder: str, image_bytes: bytes) -> None:
    def proc_block(parent):
        for p in parent.paragraphs:
            while placeholder in p.text:
                _replace_paragraph_placeholder(p, placeholder, image_bytes)
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    proc_block(cell)

    proc_block(doc)
    for sec in doc.sections:
        proc_block(sec.header)
        proc_block(sec.footer)


def _replace_paragraph_text_placeholder(paragraph, placeholder: str, text_value: str) -> None:
    before, _, after = paragraph.text.partition(placeholder)
    p_elm = paragraph._element
    for child in list(p_elm):
        p_elm.remove(child)
    if before:
        paragraph.add_run(before)
    paragraph.add_run(text_value)
    if after:
        paragraph.add_run(after)


def _walk_unirus_replace_text(doc, placeholder: str, text_value: str) -> None:
    if not text_value:
        return

    def proc_block(parent):
        for p in parent.paragraphs:
            while placeholder in p.text:
                _replace_paragraph_text_placeholder(p, placeholder, text_value)
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    proc_block(cell)

    proc_block(doc)
    for sec in doc.sections:
        proc_block(sec.header)
        proc_block(sec.footer)


def _unirus_render_error_message(code: str) -> str:
    return {
        "no_docx_lib": "На сервере не установлен пакет python-docx.",
        "template_missing": "Шаблон Unirus не найден на сервере.",
        "template_read": "Не удалось прочитать шаблон Unirus.",
        "barcode_fetch": "Не удалось загрузить изображение штрих-кода (tec-it).",
        "docx_save": "Не удалось сформировать документ Word.",
        "docx_article_patch": "Не удалось подставить артикулы в шаблон Unirus.",
    }.get(code, "Ошибка генерации паллетного листа.")


def _unirus_article_triplet_preserve_runs_sub(m: re.Match, esc: str) -> str:
    """
    Оставляет три <w:r>: крайние с пустым <w:t></w:t>, средний — артикул с исходным w:rPr
    (подчёркивание и шрифт остаются у среднего run, вёрстка документа не «схлопывается»).
    """
    apr2 = m.group("apr2") or ""
    return (
        f'{m.group("a1")}{m.group("apr1")}<w:t></w:t></w:r>'
        f'{m.group("a2")}{apr2}<w:t>{esc}</w:t></w:r>'
        f'{m.group("a3")}{m.group("apr3")}<w:t></w:t></w:r>'
    )


def _zip_repack_same_meta(zin: zipfile.ZipFile, file_data: dict[str, bytes]) -> bytes:
    """Пересобирает docx, сохраняя compress_type и прочие поля ZipInfo (меньше сюрпризов для Word)."""
    outbuf = io.BytesIO()
    zout = zipfile.ZipFile(outbuf, "w")
    try:
        for zi in zin.infolist():
            name = zi.filename
            data = file_data.get(name, zin.read(name))
            new_i = zipfile.ZipInfo(filename=name, date_time=zi.date_time)
            new_i.compress_type = zi.compress_type
            new_i.external_attr = zi.external_attr
            new_i.create_system = zi.create_system
            new_i.comment = zi.comment
            new_i.extra = zi.extra
            zout.writestr(new_i, data)
    finally:
        zout.close()
    return outbuf.getvalue()


def _patch_unirus_docx_article_bytes(docx_bytes: bytes, article: str) -> tuple[bytes | None, str]:
    """В word/document.xml подставляет артикул в тройку run ### + Артикул + ### без удаления run'ов."""
    esc = _xml_escape_text(str(article or ""), {"'": "&apos;"})

    def _sub(m: re.Match) -> str:
        return _unirus_article_triplet_preserve_runs_sub(m, esc)

    zin = None
    try:
        zin = zipfile.ZipFile(io.BytesIO(docx_bytes), "r")
        xml = zin.read("word/document.xml").decode("utf-8")
        xml2 = _UNIRUS_TRIPLET_ARTICLE_RE.sub(_sub, xml)
        return _zip_repack_same_meta(zin, {"word/document.xml": xml2.encode("utf-8")}), ""
    except Exception:
        return None, "docx_article_patch"
    finally:
        if zin is not None:
            zin.close()


def render_unirus_docx_with_pallet_barcode(
    pallet_number: str,
    article: str = "",
    extra: dict | None = None,
) -> tuple[bytes | None, str]:
    """
    Возвращает (байты .docx, код ошибки). Пустой код — успех.
    Подстановки по макросу Word: даты, ряды, итоги, партия, три Code128 (коробки, паллета, сложный).
    extra — опциональные поля из JSON (manufacturing_date_raw, expiry_date_raw, …).
    Если номер паллеты пустой, штрих-код паллеты не вставляется (как в макросе).
    """
    if not HAVE_DOCX:
        return None, "no_docx_lib"
    tpl = _unirus_template_path()
    if not tpl.is_file():
        return None, "template_missing"
    tpl_bytes = tpl.read_bytes()
    patched_tpl, perr = _patch_unirus_docx_article_bytes(tpl_bytes, article)
    if perr:
        return None, perr
    try:
        doc = DocxDocument(io.BytesIO(patched_tpl))
    except Exception:
        return None, "template_read"

    text_map, complex_data, total_boxes_str = _compute_unirus_text_replacements(
        article, extra
    )
    for ph, val in text_map.items():
        _walk_unirus_replace_text(doc, ph, val)

    if total_boxes_str:
        try:
            img_boxes = fetch_barcode_png(total_boxes_str)
        except (urllib.error.URLError, TimeoutError, OSError):
            return None, "barcode_fetch"
        _walk_unirus_replace_barcode(doc, UNIRUS_BOX_BARCODE_PLACEHOLDER, img_boxes)

    pallet_bc = build_pallet_barcode_data(pallet_number)
    if pallet_bc:
        try:
            img_pallet = fetch_barcode_png(pallet_bc)
        except (urllib.error.URLError, TimeoutError, OSError):
            return None, "barcode_fetch"
        _walk_unirus_replace_barcode(doc, UNIRUS_PALLET_BARCODE_PLACEHOLDER, img_pallet)

    if complex_data:
        try:
            img_complex = fetch_barcode_png(complex_data)
        except (urllib.error.URLError, TimeoutError, OSError):
            return None, "barcode_fetch"
        _walk_unirus_replace_barcode(doc, UNIRUS_COMPLEX_BARCODE_PLACEHOLDER, img_complex)

    out = io.BytesIO()
    try:
        doc.save(out)
    except Exception:
        return None, "docx_save"
    return out.getvalue(), ""


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
                            (item.get("name") or "").strip(),
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
        name = _normalize_str(raw.get("name", ""))
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
            "SELECT id, assembled_percent, extra_info FROM orders WHERE id = ?",
            (order_id,),
        ).fetchone()
        if not row:
            con.close()
            return {"error": "not_found", "message": "Заказ не найден."}
        apct = max(0, min(100, int(row["assembled_percent"] or 0)))
        xinfo = row["extra_info"] or ""
        cur.execute(
            """
            UPDATE orders SET ship_date = ?, client = ?, names = ?, assembled_percent = ?, extra_info = ?
            WHERE id = ?
            """,
            (ship, client_n, names_summary, apct, xinfo, order_id),
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
    }


def fetch_order_detail(order_id: int):
    with DB_LOCK:
        con = get_connection()
        row = con.execute(
            """
            SELECT id, ship_date, client, assembled_percent, names, extra_info
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
                   p.max_rows AS max_rows
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
        "items": items,
    }


def delete_order(order_id: int) -> bool:
    """Удаляет заказ; позиции order_items удаляются каскадом."""
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
            SELECT id, ship_date, client, assembled_percent, names, extra_info
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
                       p.max_rows AS max_rows
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
                "items": items,
            }
        )
    return out


def fetch_nomenclature():
    with DB_LOCK:
        con = get_connection()
        rows = con.execute(
            """
            SELECT id, article, name,
                   pieces_in_box, sets_in_box, pieces_per_set, row_layout, max_rows, box_weight
            FROM products ORDER BY id
            """
        ).fetchall()
        con.close()
    out = []
    for row in rows:
        pib = int(row["pieces_in_box"] or 0)
        sib = max(1, int(row["sets_in_box"] or 1))
        pps = _computed_pieces_per_set_for_api(pib, sib)
        out.append(
            {
                "id": row["id"],
                "article": row["article"] or "",
                "name": row["name"] or "",
                "pieces_in_box": pib,
                "sets_in_box": sib,
                "pieces_per_set": pps,
                "row_layout": int(row["row_layout"] or 0),
                "max_rows": int(row["max_rows"] or 0),
                "box_weight": float(row["box_weight"] or 0),
            }
        )
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
                name.strip(),
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
    conflict = find_nomenclature_create_conflicts(article, name)
    if conflict:
        return conflict
    norm = _normalize_packaging(pieces_in_box, sets_in_box)
    if isinstance(norm, dict):
        return norm
    pib, sib, pps = norm
    article_n = _normalize_str(article)
    name_n = _normalize_str(name)
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

    def _respond_unirus_pallet_sheet(
        self, pn_str: str, article: str = "", extra: dict | None = None
    ):
        """Отдаёт один .docx: артикул в XML, поля и штрих-коды по макросу (POST extra), GET — только паллета+артикул."""
        if not HAVE_DOCX:
            self._send_json(
                503,
                {
                    "error": "no_docx_lib",
                    "message": _unirus_render_error_message("no_docx_lib"),
                },
            )
            return
        blob, err = render_unirus_docx_with_pallet_barcode(pn_str, article, extra)
        if err:
            status_map = {
                "template_missing": 404,
                "barcode_fetch": 502,
                "no_docx_lib": 503,
                "docx_article_patch": 500,
            }
            status = status_map.get(err, 500)
            self._send_json(
                status,
                {"error": err, "message": _unirus_render_error_message(err)},
            )
            return
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        self.send_header("Content-Length", str(len(blob)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(blob)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if _is_unirus_pallet_sheet_path(path):
            qs = parse_qs(parsed.query)
            vals = qs.get("pallet_number", [""])
            raw_v = vals[0] if vals else ""
            pn_str = str(raw_v).strip()
            av = qs.get("article", [""])
            article = str(av[0]).strip() if av and av[0] is not None else ""
            self._respond_unirus_pallet_sheet(pn_str, article)
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
        if _is_unirus_template_path(path):
            tpl = _unirus_template_path()
            if not tpl.is_file():
                self._send_json(404, {"error": "Not found", "message": "Шаблон Unirus не найден."})
                return
            data = tpl.read_bytes()
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "public, max-age=3600")
            self.end_headers()
            self.wfile.write(data)
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
        if _is_unirus_pallet_sheet_path(path):
            try:
                body = self._read_json_body()
            except json.JSONDecodeError:
                self._send_json(400, {"error": "Invalid JSON"})
                return
            raw_pn = body.get("pallet_number", "")
            if raw_pn is not None and not isinstance(raw_pn, (str, int, float)):
                self._send_json(
                    400,
                    {
                        "error": "validation",
                        "message": "Поле pallet_number должно быть строкой или числом.",
                    },
                )
                return
            pn_str = str(raw_pn).strip() if raw_pn is not None else ""
            raw_art = body.get("article", "")
            article = str(raw_art).strip() if raw_art is not None else ""
            extra = _unirus_extra_from_body(body)
            self._respond_unirus_pallet_sheet(pn_str, article, extra)
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
    init_db()
    server = ThreadingHTTPServer((HOST, PORT), ApiHandler)
    print(f"API server listening on http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()

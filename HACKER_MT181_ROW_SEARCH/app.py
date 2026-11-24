
import os
import csv
import json
import re
import time
import sqlite3
import threading
import zipfile
import hashlib
import sys
import webbrowser
from collections import deque
from decimal import Decimal, InvalidOperation
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from datetime import datetime
from google_sync import GoogleWorkspace

try:
    from lxml import etree
except ImportError:
    etree = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import xlrd
except ImportError:
    xlrd = None

try:
    from pyxlsb import open_workbook as open_xlsb
except ImportError:
    open_xlsb = None

try:
    import docx
except ImportError:
    docx = None

try:
    import pypdf
except ImportError:
    pypdf = None

try:
    from charset_normalizer import from_bytes
except ImportError:
    from_bytes = None


APP_NAME = "HACKER M-T181"
DEFAULT_DATA_FOLDER = "data"
DB_NAME = "hacker_mt181_rows.db"

SUPPORTED = {
    ".xlsx",
    ".xlsm",
    ".xls",
    ".xlsb",
    ".csv",
    ".tsv",
    ".docx",
    ".doc",
    ".pdf",
    ".txt",
    ".eml", ".pptx", ".odt", ".ods", ".rtf", ".html", ".htm",
    ".md", ".log", ".json", ".jsonl", ".xml", ".ini", ".zip",
}
PC_DOCUMENTS = {
    ".xlsx", ".xlsm", ".xls", ".xlsb", ".csv", ".tsv", ".ods",
    ".docx", ".doc", ".odt", ".rtf", ".pdf", ".pptx", ".txt", ".md",
    ".html", ".htm", ".eml", ".zip",
}
PC_EXCLUDED_FOLDERS = {
    "windows", "program files", "program files (x86)", "programdata",
    "$recycle.bin", "system volume information", "recovery", "windowsapps",
    "appdata", "temp", "tmp", "cache", "caches", ".cache", "__pycache__",
    "node_modules", ".git", "msocache", "perflogs", "system32", "syswow64",
}

PAGE_SIZE = 500
INDEX_VERSION = "13.0-DUPLICATE-FILES"
DB_BATCH = 5000


# ------------------------------------------------------------
# BASIC TEXT FUNCTIONS
# ------------------------------------------------------------
def clean(value):
    if value is None:
        return ""

    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).replace("\x00", " ").strip()


def text_encoding(path):
    with open(path, "rb") as stream:
        sample = stream.read(65536)
    if sample.startswith((b"\xff\xfe", b"\xfe\xff")):
        return "utf-16"
    if sample.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    try:
        sample.decode("utf-8")
        return "utf-8-sig"
    except UnicodeDecodeError:
        pass
    if from_bytes is not None:
        guess = from_bytes(sample).best()
        if guess is not None:
            return guess.encoding
    return "cp1252"


def forms(value):
    text = clean(value).casefold()
    compact = re.sub(r"[\s,._/\\()\-]+", "", text)
    return text, compact


def cell_matches(value, query):
    q, qc = forms(query)
    v, vc = forms(value)

    if not q:
        return False

    return q in v or (qc and qc in vc)


def numeric_match_rank(cells, query):
    """Longest contiguous part of the typed digits found in one cell."""
    if any(cell_matches(value, query) for _column, value in cells):
        return len(query)
    for length in range(len(query) - 1, 3, -1):
        parts = {query[pos:pos + length] for pos in range(len(query) - length + 1)}
        if any(cell_matches(value, part) for _column, value in cells for part in parts):
            return length
    return 0


def files_under(folder):
    folder = Path(folder)

    if not folder.exists():
        return []

    return sorted(
        (
            p
            for p in folder.rglob("*")
            if p.is_file()
            and not p.name.startswith("~$")
            and not p.name.endswith(".partial")
            and ".openai-download-" not in p.name
        ),
        key=lambda p: str(p).casefold(),
    )


def local_pc_drives():
    """Enumerate locally attached Windows disks; network and optical drives are excluded."""
    if os.name != "nt":
        return []
    import ctypes
    kernel = ctypes.windll.kernel32
    mask = kernel.GetLogicalDrives()
    drives = []
    for n in range(26):
        if mask & (1 << n):
            root = f"{chr(65 + n)}:\\"
            # DRIVE_REMOVABLE=2, DRIVE_FIXED=3 (USB storage is also local).
            if kernel.GetDriveTypeW(root) in (2, 3):
                drives.append(Path(root))
    return drives


def pc_files_under(roots, on_problem=None, cancel=None, exclude_dirs=()):
    """Discover readable supported files on each local drive without following junctions."""
    excluded = {os.path.normcase(os.path.abspath(str(p))) for p in exclude_dirs}
    found = []
    seen = set()
    folders = 0
    for root in roots:
        stack = [Path(root)]
        while stack:
            if cancel and cancel.is_set():
                return found
            directory = stack.pop()
            key = os.path.normcase(os.path.abspath(str(directory)))
            if key in seen or key in excluded:
                continue
            seen.add(key)
            try:
                with os.scandir(directory) as entries:
                    for item in entries:
                        if cancel and cancel.is_set():
                            return found
                        try:
                            if item.is_symlink() or (item.stat(follow_symlinks=False).st_file_attributes & 0x400
                                                    if os.name == "nt" else False):
                                continue  # Windows reparse points can leave the drive or loop.
                            if item.is_dir(follow_symlinks=False):
                                if item.name.casefold() not in PC_EXCLUDED_FOLDERS:
                                    stack.append(Path(item.path))
                            elif (item.is_file(follow_symlinks=False)
                                  and Path(item.name).suffix.casefold() in PC_DOCUMENTS
                                  and not item.name.startswith("~$")
                                  and not item.name.endswith(".partial")):
                                found.append(Path(item.path))
                        except OSError as exc:
                            if on_problem:
                                on_problem(item.path, exc)
            except OSError as exc:
                if on_problem:
                    on_problem(directory, exc)
            folders += 1
            if on_problem and folders % 300 == 0:
                on_problem(None, (folders, len(found)))
    return found


# ------------------------------------------------------------
# FAST XLSX XML READER
# ------------------------------------------------------------
def col_letters(cell_ref):
    """
    Convert an Excel cell reference such as:
        A1    -> A
        L500  -> L
        AA77  -> AA

    Avoids openpyxl objects in the hot path.
    """
    letters = []

    for ch in cell_ref:
        if "A" <= ch <= "Z" or "a" <= ch <= "z":
            letters.append(ch.upper())
        else:
            break

    return "".join(letters) or "?"


def local_name(tag):
    if not isinstance(tag, str):
        return ""

    return tag.rsplit("}", 1)[-1]


def read_shared_strings(zf):
    """
    Read xl/sharedStrings.xml once.

    Rich-text <si> values are joined from all descendant text nodes.
    """
    try:
        raw = zf.open("xl/sharedStrings.xml")
    except KeyError:
        return []

    shared = []

    try:
        if etree is None:
            # Fallback for environments without lxml.
            import xml.etree.ElementTree as ET

            for event, elem in ET.iterparse(raw, events=("end",)):
                if local_name(elem.tag) == "si":
                    shared.append(
                        "".join(elem.itertext())
                    )
                    elem.clear()

            return shared

        context = etree.iterparse(
            raw,
            events=("end",),
            tag="{*}si",
            recover=False,
            huge_tree=True,
        )

        for _event, si in context:
            shared.append(
                "".join(
                    si.itertext()
                )
            )
            si.clear()

        return shared

    finally:
        raw.close()


def workbook_sheets(zf):
    """
    Returns:
        [(display_sheet_name, zip_member_path), ...]

    Uses workbook.xml + workbook.xml.rels so sheet names are accurate.
    """
    import xml.etree.ElementTree as ET

    workbook_root = ET.fromstring(
        zf.read("xl/workbook.xml")
    )

    rels = {}

    try:
        rels_root = ET.fromstring(
            zf.read("xl/_rels/workbook.xml.rels")
        )

        for rel in rels_root:
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target", "")

            if rid:
                if target.startswith("/"):
                    target = target.lstrip("/")

                elif not target.startswith("xl/"):
                    target = "xl/" + target

                rels[rid] = target

    except Exception:
        rels = {}

    sheets = []

    for sheet in workbook_root.iter():
        if local_name(sheet.tag) != "sheet":
            continue

        name = sheet.attrib.get("name", "Sheet")
        rid = (
            sheet.attrib.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            or sheet.attrib.get("r:id")
        )

        target = rels.get(
            rid,
            "",
        )

        if target:
            sheets.append(
                (name, target)
            )

    # Extremely defensive fallback.
    if not sheets:
        for member in sorted(
            x
            for x in zf.namelist()
            if x.startswith("xl/worksheets/")
            and x.endswith(".xml")
        ):
            sheets.append(
                (
                    Path(member).stem,
                    member,
                )
            )

    return sheets


def workbook_zero_pads(zf):
    """Map Excel style IDs to simple zero-fill widths, e.g. 00000 -> 5."""
    import xml.etree.ElementTree as ET
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except (KeyError, ET.ParseError):
        return {}
    formats = {}
    styles = {}
    for node in root:
        if local_name(node.tag) == "numFmts":
            for fmt in node:
                code = fmt.attrib.get("formatCode", "").split(";", 1)[0]
                if re.fullmatch(r"0{2,}", code):
                    formats[fmt.attrib.get("numFmtId")] = len(code)
        elif local_name(node.tag) == "cellXfs":
            for index, xf in enumerate(node):
                width = formats.get(xf.attrib.get("numFmtId"))
                if width:
                    styles[str(index)] = width
    return styles


def excel_number_text(raw, zero_width=0):
    """Expand scientific notation; honor Excel's plain 00000 formats."""
    try:
        number = Decimal(raw)
        if not number.is_finite():
            return raw
        value = format(number, "f") if "e" in raw.casefold() else raw
        if zero_width and number == number.to_integral_value():
            integer = str(int(number))
            return ("-" if integer.startswith("-") else "") + integer.lstrip("-").zfill(zero_width)
        return value
    except (InvalidOperation, ValueError):
        return raw


def read_xlsx_fast(path):
    """
    HIGH-SPEED XLSX parser.

    It does not construct openpyxl Cell objects.
    It reads the underlying XML directly and yields complete rows.

    This is substantially lighter than iterating every openpyxl cell.
    """
    if etree is None:
        # If lxml is unavailable, use the old reliable fallback.
        yield from read_xlsx_openpyxl(path)
        return

    with zipfile.ZipFile(
        path,
        "r",
    ) as zf:
        shared = read_shared_strings(
            zf
        )
        zero_pads = workbook_zero_pads(zf)

        sheets = workbook_sheets(
            zf
        )

        for sheet_name, member in sheets:
            with zf.open(member) as raw:
                context = etree.iterparse(
                    raw,
                    events=("end",),
                    tag="{*}row",
                    recover=False,
                    huge_tree=True,
                )

                for _event, row_elem in context:
                    row_number = int(
                        row_elem.attrib.get(
                            "r",
                            "0",
                        )
                    )

                    cells = []

                    for cell in row_elem:
                        if local_name(cell.tag) != "c":
                            continue

                        ref = cell.attrib.get(
                            "r",
                            "",
                        )

                        col = col_letters(
                            ref
                        )

                        ctype = cell.attrib.get(
                            "t",
                            "",
                        )

                        value = ""

                        if ctype == "inlineStr":
                            value = "".join(
                                cell.itertext()
                            )

                        else:
                            v_elem = None

                            for child in cell:
                                if local_name(
                                    child.tag
                                ) == "v":
                                    v_elem = child
                                    break

                            if v_elem is not None:
                                raw_value = (
                                    v_elem.text
                                    or ""
                                )

                                if ctype == "s":
                                    try:
                                        idx = int(
                                            raw_value
                                        )
                                        if 0 <= idx < len(
                                            shared
                                        ):
                                            value = shared[idx]
                                        else:
                                            value = raw_value
                                    except Exception:
                                        value = raw_value

                                elif ctype == "b":
                                    value = (
                                        "TRUE"
                                        if raw_value == "1"
                                        else "FALSE"
                                    )

                                else:
                                    value = excel_number_text(
                                        raw_value, zero_pads.get(cell.attrib.get("s", ""), 0)
                                    ) if ctype in {"", "n"} else raw_value

                            # Formula cells may have <f> without a useful
                            # cached <v>. Keep formula text searchable too.
                            if not value:
                                for child in cell:
                                    if local_name(
                                        child.tag
                                    ) == "f":
                                        value = (
                                            "="
                                            + "".join(
                                                child.itertext()
                                            )
                                        )
                                        break

                        value = clean(
                            value
                        )

                        if value:
                            cells.append(
                                (
                                    col,
                                    value,
                                )
                            )

                    if cells:
                        yield (
                            sheet_name,
                            row_number,
                            cells,
                        )

                    row_elem.clear()

                    # Keep the XML parser memory small.
                    while row_elem.getprevious() is not None:
                        del row_elem.getparent()[0]


def read_xlsx_openpyxl(path):
    """
    Reliable fallback when lxml is unavailable.
    """
    if openpyxl is None:
        raise RuntimeError(
            "openpyxl is missing."
        )

    wb = openpyxl.load_workbook(
        path,
        read_only=True,
        data_only=False,
        keep_links=False,
    )

    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            for row_number, row in enumerate(
                ws.iter_rows(
                    values_only=True
                ),
                start=1,
            ):
                cells = []

                for col_number, value in enumerate(
                    row,
                    start=1,
                ):
                    value = clean(
                        value
                    )

                    if value:
                        cells.append(
                            (
                                openpyxl.utils.get_column_letter(
                                    col_number
                                ),
                                value,
                            )
                        )

                if cells:
                    yield (
                        sheet_name,
                        row_number,
                        cells,
                    )
    finally:
        wb.close()


# ------------------------------------------------------------
# OTHER FILE READERS
# ------------------------------------------------------------
def read_xls(path):
    if xlrd is None:
        raise RuntimeError("Install xlrd for .xls files.")
    book = xlrd.open_workbook(str(path), on_demand=True)
    try:
        for sheet_name in book.sheet_names():
            sheet = book.sheet_by_name(sheet_name)
            for row_index in range(sheet.nrows):
                cells = []
                for col_number, cell in enumerate(sheet.row(row_index), start=1):
                    value = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            value = xlrd.xldate_as_datetime(value, book.datemode).isoformat(sep=" ")
                        except ValueError:
                            pass
                    value = clean(value)
                    if value:
                        col = openpyxl.utils.get_column_letter(col_number) if openpyxl else f"Col_{col_number}"
                        cells.append((col, value))
                if cells:
                    yield (sheet_name, row_index + 1, cells)
            book.unload_sheet(sheet_name)
    finally:
        book.release_resources()


def read_csv(path):
    with open(
        path,
        "r",
        encoding=text_encoding(path),
        errors="replace",
        newline="",
    ) as f:
        delimiter = "\t" if path.suffix.casefold() == ".tsv" else ","
        if path.suffix.casefold() == ".csv":
            sample = f.read(8192)
            f.seek(0)
            try:
                delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
            except csv.Error:
                pass
        reader = csv.reader(f, delimiter=delimiter)

        for row_number, row in enumerate(
            reader,
            start=1,
        ):
            cells = []

            for col_number, value in enumerate(
                row,
                start=1,
            ):
                value = clean(
                    value
                )

                if value:
                    col = (
                        openpyxl.utils.get_column_letter(
                            col_number
                        )
                        if openpyxl
                        else f"Col_{col_number}"
                    )

                    cells.append(
                        (
                            col,
                            value,
                        )
                    )

            if cells:
                yield (
                    "CSV",
                    row_number,
                    cells,
                )


def read_xlsb(path):
    if open_xlsb is None:
        raise RuntimeError("Install pyxlsb for .xlsb files.")
    with open_xlsb(str(path)) as workbook:
        for sheet_name in workbook.sheets:
            with workbook.get_sheet(sheet_name) as sheet:
                for row in sheet.rows(sparse=True):
                    cells = []
                    for cell in row:
                        value = clean(cell.v)
                        if value:
                            col = (openpyxl.utils.get_column_letter(cell.c + 1)
                                   if openpyxl else f"Col_{cell.c+1}")
                            cells.append((col, value))
                    if cells:
                        yield sheet_name, row[0].r + 1, cells


def read_docx(path):
    if docx is None:
        raise RuntimeError(
            "python-docx is missing."
        )

    document = docx.Document(
        path
    )

    for n, paragraph in enumerate(
        document.paragraphs,
        start=1,
    ):
        value = clean(
            paragraph.text
        )

        if value:
            yield (
                "Paragraph",
                n,
                [
                    (
                        "Text",
                        value,
                    )
                ],
            )

    for table_no, table in enumerate(
        document.tables,
        start=1,
    ):
        for row_no, row in enumerate(
            table.rows,
            start=1,
        ):
            cells = []

            for col_no, cell in enumerate(
                row.cells,
                start=1,
            ):
                value = clean(
                    cell.text
                )

                if value:
                    cells.append(
                        (
                            f"Col_{col_no}",
                            value,
                        )
                    )

            if cells:
                yield (
                    f"Table {table_no}",
                    row_no,
                    cells,
                )

            for col_no, cell in enumerate(row.cells, 1):
                for nested_no, nested in enumerate(cell.tables, 1):
                    for nested_row_no, nested_row in enumerate(nested.rows, 1):
                        nested_cells = [
                            (f"Col_{i}", clean(nc.text))
                            for i, nc in enumerate(nested_row.cells, 1)
                            if clean(nc.text)
                        ]
                        if nested_cells:
                            yield (
                                f"Table {table_no} / cell {col_no} / nested {nested_no}",
                                nested_row_no, nested_cells,
                            )

    for section_no, section in enumerate(document.sections, 1):
        for label, region in (("Header", section.header), ("Footer", section.footer)):
            for n, paragraph in enumerate(region.paragraphs, 1):
                if paragraph.text.strip():
                    yield f"{label} {section_no}", n, [("Text", paragraph.text.strip())]
            for table_no, table in enumerate(region.tables, 1):
                for row_no, row in enumerate(table.rows, 1):
                    cells = [(f"Col_{i}", cell.text) for i, cell in enumerate(row.cells, 1) if cell.text]
                    if cells:
                        yield f"{label} {section_no} / table {table_no}", row_no, cells


def read_pdf(path):
    if pypdf is None:
        raise RuntimeError(
            "pypdf is missing."
        )

    reader = pypdf.PdfReader(
        path
    )

    for page_no, page in enumerate(
        reader.pages,
        start=1,
    ):
        text = page.extract_text() or ""
        if not text.strip() and page.images:
            raise RuntimeError(
                f"PDF page {page_no} is image-only; OCR is required before numeric matches can be guaranteed"
            )

        for line_no, line in enumerate(
            text.splitlines(),
            start=1,
        ):
            value = clean(
                line
            )

            if value:
                yield (
                    f"Page {page_no}",
                    line_no,
                    [
                        (
                            "Text",
                            value,
                        )
                    ],
                )


def read_txt(path):
    with open(
        path,
        "r",
        encoding=text_encoding(path),
        errors="replace",
    ) as f:
        for line_no, line in enumerate(
            f,
            start=1,
        ):
            value = clean(
                line
            )

            if value:
                yield (
                    "TXT",
                    line_no,
                    [
                        (
                            "Text",
                            value,
                        )
                    ],
                )


def read_rows(path):
    ext = path.suffix.casefold()

    if ext in {
        ".xlsx",
        ".xlsm",
    }:
        yield from read_xlsx_fast(
            path
        )

    elif ext == ".xls":
        yield from read_xls(
            path
        )

    elif ext == ".xlsb":
        yield from read_xlsb(path)

    elif ext in {".csv", ".tsv"}:
        yield from read_csv(
            path
        )

    elif ext == ".docx":
        yield from read_docx(
            path
        )

    elif ext == ".pdf":
        yield from read_pdf(
            path
        )

    elif ext == ".txt":
        yield from read_txt(
            path
        )

    else:
        from extra_readers import read_extra_rows
        yield from read_extra_rows(path, read_rows)


def full_row_text(cells):
    return " | ".join(
        f"[{col}]: {value}"
        for col, value in cells
    )


def search_text(cells):
    """One searchable string per row. No trigram/FTS index is built."""
    normal_parts = []
    compact_parts = []

    for _, value in cells:
        if value:
            text, compact = forms(value)
            if text:
                normal_parts.append(text)
            if compact:
                compact_parts.append(compact)

    return " ".join(normal_parts) + " " + " ".join(compact_parts)


def chars_for_row(cells):
    text = "".join(
        forms(value)[1]
        for _, value in cells
        if value
    )

    return {
        char
        for char in text
        if char.isalnum()
    }


# ------------------------------------------------------------
# DATABASE INDEX
# ------------------------------------------------------------
class FastIndex:
    """Fast disposable SQLite index optimized for very large Excel files.

    V9 deliberately removes SQLite FTS5/trigram indexing. FTS5 is excellent for
    repeated substring queries, but for this application it adds a large amount
    of write/index-build work. Instead we keep one compact lowercase search_blob
    per row and use SQLite's native INSTR() scan. The scan is performed in C and
    avoids Python row-by-row comparisons until matching rows are found.

    Important:
      * Row 1 is searchable.
      * All sheets and subfolders are included.
      * Source files are read only.
      * The database can always be rebuilt because it is only a cache/index.
    """

    def __init__(self, db_path):
        self.db_path = str(db_path)
        self.lock = threading.RLock()
        self._scan_in_progress = False

        self.conn = sqlite3.connect(
            self.db_path,
            check_same_thread=False,
            timeout=180,
        )
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA synchronous=NORMAL")
        self.conn.execute("PRAGMA temp_store=MEMORY")
        self.conn.execute("PRAGMA cache_size=-100000")
        self.conn.execute("PRAGMA mmap_size=1073741824")
        self.conn.execute("PRAGMA foreign_keys=ON")
        self.setup()

    def setup(self):
        with self.lock:
            c = self.conn.cursor()
            c.execute("""
                CREATE TABLE IF NOT EXISTS meta(
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
            """)
            c.execute("""
                CREATE TABLE IF NOT EXISTS files(
                    id INTEGER PRIMARY KEY,
                    path TEXT UNIQUE NOT NULL,
                    mtime_ns INTEGER NOT NULL,
                    file_size INTEGER NOT NULL,
                    content_sha256 TEXT NOT NULL DEFAULT '',
                    row_count INTEGER NOT NULL DEFAULT 0,
                    indexed_ok INTEGER NOT NULL DEFAULT 0,
                    last_error TEXT DEFAULT ''
                )
            """)
            c.execute("""
                CREATE TABLE IF NOT EXISTS rows(
                    id INTEGER PRIMARY KEY,
                    file_id INTEGER NOT NULL,
                    sheet_name TEXT NOT NULL,
                    row_number INTEGER NOT NULL,
                    search_blob TEXT NOT NULL,
                    cells_json TEXT NOT NULL,
                    FOREIGN KEY(file_id) REFERENCES files(id) ON DELETE CASCADE
                )
            """)
            c.execute("CREATE INDEX IF NOT EXISTS idx_files_ok ON files(indexed_ok)")
            c.execute("CREATE INDEX IF NOT EXISTS idx_rows_file ON rows(file_id)")

            c.execute("SELECT value FROM meta WHERE key='version'")
            current = c.fetchone()
            if not current or current[0] != INDEX_VERSION:
                # Index is disposable. Rebuild cleanly when the format changes.
                try:
                    c.execute("DROP TABLE IF EXISTS rows")
                    c.execute("DROP TABLE IF EXISTS files")
                    c.execute("DROP TABLE IF EXISTS meta")
                    c.execute("CREATE TABLE meta(key TEXT PRIMARY KEY, value TEXT)")
                    c.execute("""
                        CREATE TABLE files(
                            id INTEGER PRIMARY KEY,
                            path TEXT UNIQUE NOT NULL,
                            mtime_ns INTEGER NOT NULL,
                            file_size INTEGER NOT NULL,
                            content_sha256 TEXT NOT NULL DEFAULT '',
                            row_count INTEGER NOT NULL DEFAULT 0,
                            indexed_ok INTEGER NOT NULL DEFAULT 0,
                            last_error TEXT DEFAULT ''
                        )
                    """)
                    c.execute("""
                        CREATE TABLE rows(
                            id INTEGER PRIMARY KEY,
                            file_id INTEGER NOT NULL,
                            sheet_name TEXT NOT NULL,
                            row_number INTEGER NOT NULL,
                            search_blob TEXT NOT NULL,
                            cells_json TEXT NOT NULL,
                            FOREIGN KEY(file_id) REFERENCES files(id) ON DELETE CASCADE
                        )
                    """)
                    c.execute("CREATE INDEX idx_files_ok ON files(indexed_ok)")
                    c.execute("CREATE INDEX idx_rows_file ON rows(file_id)")
                    c.execute(
                        "INSERT INTO meta(key,value) VALUES('version',?)",
                        (INDEX_VERSION,),
                    )
                except Exception:
                    self.conn.rollback()
                    raise
            c.execute("""
                CREATE TABLE IF NOT EXISTS search_cache(
                    query TEXT NOT NULL,
                    after_id INTEGER NOT NULL,
                    page_size INTEGER NOT NULL,
                    payload TEXT NOT NULL,
                    has_more INTEGER NOT NULL,
                    saved_at TEXT NOT NULL,
                    PRIMARY KEY(query,after_id,page_size)
                )
            """)
            c.execute("""
                CREATE TABLE IF NOT EXISTS searched_numbers(
                    query TEXT PRIMARY KEY,
                    first_at TEXT NOT NULL,
                    last_at TEXT NOT NULL,
                    times INTEGER NOT NULL,
                    last_matches TEXT NOT NULL
                )
            """)
            self.conn.commit()

    def previous_query(self, query):
        with self.lock:
            row = self.conn.execute(
                "SELECT first_at,last_at,times,last_matches FROM searched_numbers WHERE query=?",
                (query,),
            ).fetchone()
        return row

    def list_checked_numbers(self, offset=0, limit=300, exact=""):
        with self.lock:
            if exact:
                rows = self.conn.execute(
                    "SELECT query,first_at,last_at,times,last_matches FROM searched_numbers WHERE query=?",
                    (exact,),
                ).fetchall()
                return rows, len(rows)
            total = self.conn.execute("SELECT COUNT(*) FROM searched_numbers").fetchone()[0]
            rows = self.conn.execute(
                "SELECT query,first_at,last_at,times,last_matches FROM searched_numbers "
                "ORDER BY last_at DESC,query DESC LIMIT ? OFFSET ?",
                (limit, offset),
            ).fetchall()
        return rows, total

    def record_query(self, query, matches):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.lock:
            self.conn.execute("""
                INSERT INTO searched_numbers(query,first_at,last_at,times,last_matches)
                VALUES(?,?,?,1,?)
                ON CONFLICT(query) DO UPDATE SET last_at=excluded.last_at,
                times=times+1,last_matches=excluded.last_matches
            """, (query, now, now, matches))
            self.conn.commit()
        return self.previous_query(query)

    def import_history(self, history_path):
        if not Path(history_path).is_file():
            return
        with self.lock:
            if self.conn.execute("SELECT 1 FROM searched_numbers LIMIT 1").fetchone():
                return
            with Path(history_path).open(encoding="utf-8") as stream:
                for line in stream:
                    try:
                        item = json.loads(line)
                        query = item["query"]
                        if not query.isascii() or not query.isdigit():
                            continue
                        when = item["time"]
                        self.conn.execute("""
                            INSERT INTO searched_numbers(query,first_at,last_at,times,last_matches)
                            VALUES(?,?,?,1,?)
                            ON CONFLICT(query) DO UPDATE SET last_at=excluded.last_at,
                            times=times+1,last_matches=excluded.last_matches
                        """, (query, when, when, str(item.get("matches", "?"))))
                    except (ValueError, KeyError, TypeError):
                        continue
            self.conn.commit()

    def close(self):
        with self.lock:
            try:
                self.conn.close()
            except Exception:
                pass

    def stats(self):
        with self.lock:
            c = self.conn.cursor()
            c.execute("SELECT COUNT(*) FROM files WHERE indexed_ok=1")
            good = c.fetchone()[0]
            c.execute("SELECT COUNT(*) FROM files")
            total = c.fetchone()[0]
            c.execute("SELECT COALESCE(SUM(row_count),0) FROM files WHERE indexed_ok=1")
            rows = c.fetchone()[0]
            c.execute("SELECT COUNT(*) FROM files WHERE indexed_ok=0")
            failed = c.fetchone()[0]
            return good, total, rows, failed

    def delete_file(self, file_id):
        with self.lock:
            c = self.conn.cursor()
            c.execute("DELETE FROM rows WHERE file_id=?", (file_id,))
            c.execute("DELETE FROM files WHERE id=?", (file_id,))
            self.conn.commit()

    def scan(self, folder, progress=None, cancel=None, pc_mode=False, exclude_dirs=()):
        with self.lock:
            self._scan_in_progress = True
            # A rescan may change any file. Never reuse a page from an older
            # snapshot while rows are being replaced.
            self.conn.execute("DELETE FROM search_cache")
            self.conn.commit()
        directory_errors = []
        if pc_mode:
            def discovery_problem(path, issue):
                if path is None:
                    if progress:
                        progress("discovery", issue[0], 0, "Local drives", issue[1])
                else:
                    directory_errors.append(f"{path}: {issue}")
                    if progress and len(directory_errors) <= 25:
                        progress("directory_error", 0, 0, path, 0, str(issue))
            files = pc_files_under(folder, discovery_problem, cancel, exclude_dirs)
        else:
            files = files_under(folder)
        if cancel and cancel.is_set():
            with self.lock:
                self._scan_in_progress = False
            good, _total, rows, failed = self.stats()
            return {"ok_files": good, "files": 0, "rows": rows,
                    "failed": failed, "errors": [], "unreadable_dirs": 0,
                    "cancelled": True}
        total_files = len(files)

        disk_paths = {str(p.resolve()) for p in files}
        with self.lock:
            c = self.conn.cursor()
            c.execute("SELECT path,id FROM files")
            db_paths = {p: fid for p, fid in c.fetchall()}

        for old_path in (set(db_paths) - disk_paths):
            self.delete_file(db_paths[old_path])

        errors = directory_errors[:25]
        started = time.perf_counter()

        for file_no, path in enumerate(files, start=1):
            if cancel and cancel.is_set():
                break

            path_str = str(path.resolve())
            try:
                st = path.stat()
                sig = (st.st_mtime_ns, st.st_size)
            except Exception as exc:
                errors.append(f"{path.name}: {exc}")
                with self.lock:
                    self.conn.execute(
                        "INSERT INTO files(path,mtime_ns,file_size,content_sha256,row_count,indexed_ok,last_error) "
                        "VALUES(?,0,0,'',0,0,?) ON CONFLICT(path) DO UPDATE SET "
                        "row_count=0,indexed_ok=0,last_error=excluded.last_error",
                        (path_str, str(exc)),
                    )
                    fid = self.conn.execute("SELECT id FROM files WHERE path=?", (path_str,)).fetchone()[0]
                    self.conn.execute("DELETE FROM rows WHERE file_id=?", (fid,))
                    self.conn.commit()
                if progress:
                    progress("error", file_no, total_files, path, 0, str(exc))
                continue

            with self.lock:
                c = self.conn.cursor()
                c.execute(
                    "SELECT id,mtime_ns,file_size,indexed_ok,row_count,content_sha256 FROM files WHERE path=?",
                    (path_str,),
                )
                cached = c.fetchone()

            digest = None
            if cached and cached[3] == 1 and (cached[1], cached[2]) == sig:
                if pc_mode:
                    if progress:
                        progress("cached", file_no, total_files, path, cached[4])
                    continue
                try:
                    with path.open("rb") as source_file:
                        digest = hashlib.file_digest(source_file, "sha256").hexdigest()
                    if digest == cached[5]:
                        if progress:
                            progress("cached", file_no, total_files, path, cached[4])
                        continue
                except Exception as exc:
                    errors.append(f"{path.name}: {exc}")
                    if progress:
                        progress("error", file_no, total_files, path, 0, str(exc))
                    continue

            if progress:
                progress("start", file_no, total_files, path, 0)

            file_id = cached[0] if cached else None
            row_count = 0

            try:
                if digest is None:
                    with path.open("rb") as source_file:
                        digest = hashlib.file_digest(source_file, "sha256").hexdigest()

                # Same bytes with a changed timestamp need no parse or copy.
                if cached and cached[3] == 1 and cached[5] == digest:
                    with self.lock:
                        self.conn.execute(
                            "UPDATE files SET mtime_ns=?,file_size=? WHERE id=?",
                            (st.st_mtime_ns, st.st_size, file_id),
                        )
                        self.conn.commit()
                    if progress:
                        progress("cached", file_no, total_files, path, cached[4])
                    continue

                with self.lock:
                    source_row = self.conn.execute(
                        "SELECT id,row_count FROM files "
                        "WHERE content_sha256=? AND indexed_ok=1 AND path<>? LIMIT 1",
                        (digest, path_str),
                    ).fetchone()
                if source_row:
                    with self.lock:
                        c = self.conn.cursor()
                        c.execute("BEGIN IMMEDIATE")
                        if file_id is None:
                            c.execute(
                                "INSERT INTO files(path,mtime_ns,file_size,content_sha256,row_count,indexed_ok) "
                                "VALUES(?,?,?,?,0,2)",
                                (path_str, st.st_mtime_ns, st.st_size, digest),
                            )
                            file_id = c.lastrowid
                        else:
                            c.execute("DELETE FROM rows WHERE file_id=?", (file_id,))
                            c.execute(
                                "UPDATE files SET mtime_ns=?,file_size=?,content_sha256=?,row_count=0,indexed_ok=2 "
                                "WHERE id=?",
                                (st.st_mtime_ns, st.st_size, digest, file_id),
                            )
                        c.execute(
                            "INSERT INTO rows(file_id,sheet_name,row_number,search_blob,cells_json) "
                            "SELECT ?,sheet_name,row_number,search_blob,cells_json FROM rows WHERE file_id=?",
                            (file_id, source_row[0]),
                        )
                        current = path.stat()
                        if (current.st_mtime_ns, current.st_size) != sig:
                            raise RuntimeError("File changed during scan; rescan after saving it.")
                        c.execute(
                            "UPDATE files SET row_count=?,indexed_ok=1,last_error='' WHERE id=?",
                            (source_row[1], file_id),
                        )
                        self.conn.commit()
                    if progress:
                        progress("copied", file_no, total_files, path, source_row[1])
                    continue

                with self.lock:
                    c = self.conn.cursor()
                    c.execute("BEGIN IMMEDIATE")

                    if file_id is not None:
                        c.execute("DELETE FROM rows WHERE file_id=?", (file_id,))
                        c.execute(
                            "UPDATE files SET mtime_ns=?,file_size=?,content_sha256=?,row_count=0,indexed_ok=2,last_error='' WHERE id=?",
                            (st.st_mtime_ns, st.st_size, digest, file_id),
                        )
                    else:
                        c.execute(
                            "INSERT INTO files(path,mtime_ns,file_size,content_sha256,row_count,indexed_ok,last_error) VALUES(?,?,?,?,?,2,'')",
                            (path_str, st.st_mtime_ns, st.st_size, digest, 0),
                        )
                        file_id = c.lastrowid

                    c.execute("SELECT COALESCE(MAX(id),0) FROM rows")
                    next_row_id = c.fetchone()[0] + 1

                rows_batch = []
                last_progress = time.perf_counter()

                for sheet_name, row_number, cells in read_rows(path):
                    if cancel and cancel.is_set():
                        raise RuntimeError("Scan cancelled.")

                    rid = next_row_id
                    next_row_id += 1
                    row_count += 1

                    # Only two persisted copies are kept for speed:
                    # searchable text + the exact cells needed for display.
                    blob = search_text(cells)
                    cells_json = json.dumps(cells, ensure_ascii=False, separators=(",", ":"))
                    rows_batch.append(
                        (rid, file_id, sheet_name, row_number, blob, cells_json)
                    )

                    if len(rows_batch) >= 10000:
                        with self.lock:
                            c = self.conn.cursor()
                            c.executemany(
                                """
                                INSERT INTO rows
                                (id,file_id,sheet_name,row_number,search_blob,cells_json)
                                VALUES(?,?,?,?,?,?)
                                """,
                                rows_batch,
                            )
                            # Commit each large batch so LIVE SEARCH can see
                            # rows as soon as they are loaded, instead of
                            # waiting for the entire large Excel file.
                            self.conn.commit()
                            self.conn.execute("BEGIN IMMEDIATE")
                        rows_batch.clear()

                    now = time.perf_counter()
                    if progress and (now - last_progress >= 0.5):
                        progress("reading", file_no, total_files, path, row_count)
                        last_progress = now

                if path.suffix.casefold() == ".pdf" and st.st_size and not row_count:
                    raise RuntimeError("PDF has no extractable text; OCR is required.")

                if rows_batch:
                    with self.lock:
                        c = self.conn.cursor()
                        c.executemany(
                            """
                            INSERT INTO rows
                            (id,file_id,sheet_name,row_number,search_blob,cells_json)
                            VALUES(?,?,?,?,?,?)
                            """,
                            rows_batch,
                        )

                current = path.stat()
                if (current.st_mtime_ns, current.st_size) != sig:
                    raise RuntimeError("File changed during scan; rescan after saving it.")

                with self.lock:
                    c = self.conn.cursor()
                    c.execute(
                        "UPDATE files SET row_count=?,indexed_ok=1,last_error='' WHERE id=?",
                        (row_count, file_id),
                    )
                    self.conn.commit()

                if progress:
                    elapsed = time.perf_counter() - started
                    progress("done", file_no, total_files, path, row_count)

            except Exception as exc:
                error_text = str(exc)
                with self.lock:
                    try:
                        self.conn.rollback()
                    except Exception:
                        pass
                    c = self.conn.cursor()
                    c.execute(
                        "INSERT INTO files(path,mtime_ns,file_size,content_sha256,row_count,indexed_ok,last_error) "
                        "VALUES(?,?,?,'',0,0,?) ON CONFLICT(path) DO UPDATE SET "
                        "mtime_ns=excluded.mtime_ns,file_size=excluded.file_size,"
                        "content_sha256='',row_count=0,indexed_ok=0,last_error=excluded.last_error",
                        (path_str, st.st_mtime_ns, st.st_size, error_text),
                    )
                    fid = c.execute("SELECT id FROM files WHERE path=?", (path_str,)).fetchone()[0]
                    c.execute("DELETE FROM rows WHERE file_id=?", (fid,))
                    self.conn.commit()

                errors.append(f"{path.name}: {error_text}")
                if progress:
                    progress("error", file_no, total_files, path, 0, error_text)

        good, total, rows, failed = self.stats()
        with self.lock:
            self._scan_in_progress = False
        return {
            "ok_files": good,
            "files": total_files,
            "rows": rows,
            "failed": max(failed, len(errors), len(directory_errors)),
            "unreadable_dirs": len(directory_errors),
            "errors": errors,
        }

    def search_page(self, query, after_id=0, page_size=PAGE_SIZE):
        if not query or len(query) > 18:
            return [], False
        q, qc = forms(query)
        if not q:
            return [], False

        # INSTR is a compiled SQLite string operation. This avoids pulling all
        # rows into Python just to test the query and avoids the heavy trigram
        # index that made large-file indexing slow.
        # indexed_ok=2 means the file is currently scanning but its already-
        # inserted rows are immediately eligible for LIVE SEARCH.
        with self.lock:
            c = self.conn.cursor()
            cached = None if self._scan_in_progress else c.execute(
                "SELECT payload,has_more FROM search_cache WHERE query=? AND after_id=? AND page_size=?",
                (query, after_id, page_size),
            ).fetchone()
            if cached:
                return json.loads(cached[0]), bool(cached[1])
            c.execute(
                """
                SELECT rows.id, files.path, rows.sheet_name, rows.row_number, rows.cells_json
                FROM rows
                CROSS JOIN files ON files.id=rows.file_id
                WHERE files.indexed_ok IN (1,2)
                  AND rows.id > ?
                  AND (instr(rows.search_blob, ?) > 0 OR instr(rows.search_blob, ?) > 0)
                ORDER BY rows.id
                """,
                (after_id, q, qc or q),
            )
            results = []
            # Inspect the actual cells; a blob can match across cell boundaries.
            # Fetch one extra real match so pagination never loses a row.
            while len(results) <= page_size:
                fetched = c.fetchmany(min(512, page_size + 1))
                if not fetched:
                    break
                for rid, path, sheet_name, row_number, cells_json in fetched:
                    cells = json.loads(cells_json)
                    matched = [
                        (col, value) for col, value in cells
                        if cell_matches(value, query)
                    ]
                    if matched:
                        results.append({
                            "id": rid, "path": path, "sheet": sheet_name,
                            "row": row_number, "cells": cells,
                            "matched": matched, "full": full_row_text(cells),
                        })
                        if len(results) > page_size:
                            break
            page = results[:page_size]
            more = len(results) > page_size
            if not self._scan_in_progress:
                c.execute(
                    "INSERT OR REPLACE INTO search_cache VALUES(?,?,?,?,?,datetime('now'))",
                    (query, after_id, page_size, json.dumps(page, ensure_ascii=False), int(more)),
                )
                c.execute("""
                    DELETE FROM search_cache WHERE rowid IN
                    (SELECT rowid FROM search_cache ORDER BY saved_at DESC,rowid DESC LIMIT -1 OFFSET 128)
                """)
                self.conn.commit()
        return page, more

    def search(self, query):
        """Return the first page for headless callers; UI uses pagination."""
        return self.search_page(query)[0]

    def search_files_page(self, query, after_id=0, page_size=PAGE_SIZE):
        """Find document filenames/paths; one result per file, including unreadable files."""
        if not query or len(query) > 18:
            return [], False
        folded = query.casefold()
        with self.lock:
            if query.isascii():
                cursor = self.conn.execute(
                    "SELECT id,path,file_size,row_count,indexed_ok,last_error FROM files "
                    "WHERE id>? AND instr(lower(path),?)>0 ORDER BY id",
                    (after_id, folded),
                )
            else:
                cursor = self.conn.execute(
                    "SELECT id,path,file_size,row_count,indexed_ok,last_error FROM files "
                    "WHERE id>? ORDER BY id", (after_id,)
                )
            found = []
            while len(found) <= page_size:
                batch = cursor.fetchmany(500)
                if not batch:
                    break
                for rid, path, size, count, ok, error in batch:
                    if folded not in path.casefold():
                        continue
                    name = Path(path).name
                    cells = [("File name", name), ("Full path", path),
                             ("Size (bytes)", str(size)), ("Indexed rows", str(count)),
                             ("Status", "Indexed" if ok == 1 else
                              "Loading" if ok == 2 else f"Unreadable: {error}")]
                    found.append({"id": rid, "path": path, "sheet": "File",
                                  "row": "—", "cells": cells,
                                  "matched": [("File name" if folded in name.casefold()
                                               else "Full path", name if folded in name.casefold()
                                               else path)], "full": full_row_text(cells)})
                    if len(found) > page_size:
                        break
            return found[:page_size], len(found) > page_size

    def near_rank_page(self, query, rank, after_id, page_size, cancelled=None):
        """Page one similarity tier by row ID without materializing the whole database."""
        cache_key = f"near:{query}:{rank}"
        with self.lock:
            cached = self.conn.execute(
                "SELECT payload,has_more FROM search_cache "
                "WHERE query=? AND after_id=? AND page_size=?",
                (cache_key, after_id, page_size),
            ).fetchone()
        if cached:
            return json.loads(cached[0]), bool(cached[1])
        parts = sorted({query[pos:pos + rank] for pos in range(len(query) - rank + 1)})
        terms = sorted({form for part in parts for form in forms(part) if form})
        predicate = " OR ".join("instr(rows.search_blob, ?) > 0" for _ in terms)
        conn = sqlite3.connect(self.db_path, timeout=180)
        try:
            if cancelled:
                conn.set_progress_handler(lambda: 1 if cancelled() else 0, 10000)
            cursor = conn.execute(
                "SELECT rows.id,files.path,rows.sheet_name,rows.row_number,rows.cells_json "
                "FROM rows CROSS JOIN files ON files.id=rows.file_id "
                "WHERE files.indexed_ok IN (1,2) AND rows.id>? AND (" + predicate + ") "
                "ORDER BY rows.id",
                (after_id, *terms),
            )
            found = []
            while len(found) <= page_size:
                if cancelled and cancelled():
                    raise InterruptedError("Search replaced by a newer query")
                batch = cursor.fetchmany(512)
                if not batch:
                    break
                for rid, path, sheet, row, payload in batch:
                    cells = json.loads(payload)
                    if numeric_match_rank(cells, query) != rank:
                        continue
                    matched = [(col, value) for col, value in cells
                               if any(cell_matches(value, part) for part in parts)]
                    found.append({"id": rid, "path": path, "sheet": sheet,
                                  "row": row, "cells": cells, "matched": matched,
                                  "full": full_row_text(cells), "rank": rank})
                    if len(found) > page_size:
                        break
            page, more = found[:page_size], len(found) > page_size
            with self.lock:
                self.conn.execute(
                    "INSERT OR REPLACE INTO search_cache VALUES(?,?,?,?,?,datetime('now'))",
                    (cache_key, after_id, page_size,
                     json.dumps(page, ensure_ascii=False), int(more)),
                )
                self.conn.execute(
                    "DELETE FROM search_cache WHERE rowid IN "
                    "(SELECT rowid FROM search_cache ORDER BY saved_at DESC,rowid DESC "
                    "LIMIT -1 OFFSET 128)"
                )
                self.conn.commit()
            return page, more
        finally:
            conn.close()

    def search_ranked_page(self, query, cursor=None, page_size=PAGE_SIZE,
                           cancelled=None, exact_ready=None):
        """Exact rows first, then 1 digit shorter down to four; paginate all tiers."""
        if not query or len(query) > 18:
            return [], False, None
        if not (query.isascii() and query.isdigit()):
            after_id = cursor[1] if cursor else 0
            page, more = self.search_page(query, after_id, page_size)
            for row in page:
                row["rank"] = len(query)
            if exact_ready:
                exact_ready(page, more)
            return page, more, (len(query), page[-1]["id"]) if more else None
        rank, after_id = cursor or (len(query), 0)
        results = []
        while rank >= 4 or rank == len(query):
            if cancelled and cancelled():
                raise InterruptedError("Search replaced by a newer query")
            remaining = page_size - len(results)
            if rank == len(query):
                page, more = self.search_page(query, after_id, remaining)
                for row in page:
                    row["rank"] = rank
                if exact_ready:
                    exact_ready(page, more)
            else:
                page, more = self.near_rank_page(query, rank, after_id, remaining, cancelled)
            results.extend(page)
            if more:
                return results, True, (rank, page[-1]["id"])
            rank -= 1
            after_id = 0
            if len(results) == page_size:
                return results, rank >= 4, (rank, 0) if rank >= 4 else None
        return results, False, None

    def export_matches(self, query, output_path, scope="data"):
        count = 0
        after_id = 0
        with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["File", "Type", "Sheet", "Row", "Matched Cell(s)",
                        "Complete Row - ALL COLUMNS"])
            while True:
                results, more = (self.search_files_page(query, after_id, 1000)
                                 if scope == "file" else self.search_page(query, after_id, 1000))
                for r in results:
                    w.writerow([
                        r["path"], Path(r["path"]).suffix.lstrip(".").upper(),
                        r["sheet"], r["row"],
                        " | ".join(f"[{c}]: {v}" for c, v in r["matched"]),
                        r["full"],
                    ])
                    count += 1
                if not more:
                    break
                after_id = results[-1]["id"]
        return count


# ------------------------------------------------------------
# GUI
# ------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root

        self.root.title(
            f"{APP_NAME} | LIVE NUMERIC SEARCH"
        )
        screen_w, screen_h = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"{min(1580, screen_w)}x{min(900, screen_h)}")
        self.root.minsize(min(950, screen_w), min(600, screen_h))
        self.root.configure(
            bg="#080d16"
        )

        self.base_dir = Path(sys.executable).resolve().parent if getattr(
            sys, "frozen", False
        ) else Path(__file__).resolve().parent

        self.folder = (
            self.base_dir
            / DEFAULT_DATA_FOLDER
        )

        self.folder.mkdir(
            parents=True,
            exist_ok=True,
        )

        self.db_path = (
            self.base_dir
            / DB_NAME
        )
        self.pc_state_path = self.base_dir / "pc_connection.json"
        try:
            self.pc_connected = (os.name == "nt" and
                json.loads(self.pc_state_path.read_text(encoding="utf-8")).get("connected") is True)
        except (OSError, ValueError, TypeError):
            self.pc_connected = False
        self.pending_pc_connect = False
        self.history_path = self.base_dir / "search_history.jsonl"
        self.log_path = self.base_dir / "activity.log"
        self.cloud = GoogleWorkspace(
            self.base_dir / "google_config", self.folder,
            notify=lambda line: self.root.after(0, self.google_progress, line),
        )

        self.index = FastIndex(
            self.base_dir / "hacker_mt181_pc_index.db" if self.pc_connected else self.db_path
        )

        self.scanning = False
        self.exporting = False
        self.last_scan_failed = 0
        self.cancel_scan = threading.Event()

        self.search_after = None
        self.search_token = 0
        self.pending_query = ""
        self.search_mode = "exact"
        self.search_scope = tk.StringVar(value="data")
        self.current_scope = "data"
        self.preview_cursor = None
        self.last_live_refresh = 0.0

        self.current_query = ""
        self.current_results = []
        self.page_starts = [0]
        self.page_index = 0
        self.has_more = False
        self.checked_offset = 0
        self.cloud_connected = False
        self.google_connecting = False
        self.cloud_syncing = False
        self.cloud_waiting = False
        self.sync_needs_scan = False
        self.last_cloud_failed = 0
        self.previous_search = None
        self.folder_snapshot = None

        self.setup_styles()
        self.setup_ui()
        self.load_history()
        self.refresh_checked_numbers()
        self.load_activity()
        self.refresh_file_panel()

        self.root.after(
            250,
            self.start_scan,
        )
        self.root.after(15000, self.check_folder_changes)
        if self.cloud.token_path.exists():
            self.root.after(1000, self.connect_google_saved)

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Treeview",
            background="#111827",
            foreground="#e5e7eb",
            fieldbackground="#111827",
            rowheight=29,
            font=("Segoe UI", 9),
        )

        style.map(
            "Treeview",
            background=[
                ("selected", "#075985")
            ],
            foreground=[
                ("selected", "#ffffff")
            ],
        )

        style.configure(
            "Treeview.Heading",
            background="#1f2937",
            foreground="#22d3ee",
            font=("Segoe UI", 9, "bold"),
        )

        style.configure(
            "Fast.Horizontal.TProgressbar",
            background="#22d3ee",
            troughcolor="#1f2937",
        )

    def setup_ui(self):
        header = tk.Frame(
            self.root,
            bg="#0f172a",
            padx=16,
            pady=11,
        )
        header.pack(
            fill="x"
        )

        tk.Label(
            header,
            text=f"⚡ {APP_NAME}",
            font=("Segoe UI", 17, "bold"),
            fg="#22d3ee",
            bg="#0f172a",
        ).pack(
            side="left"
        )

        tk.Label(
            header,
            text=" | ALL FILES · NUMBERS + TEXT",
            font=("Segoe UI", 10, "bold"),
            fg="#94a3b8",
            bg="#0f172a",
        ).pack(
            side="left",
            padx=7,
        )

        self.lbl_stats = tk.Label(
            header,
            text="Starting...",
            font=("Segoe UI", 10, "bold"),
            fg="#34d399",
            bg="#0f172a",
        )
        self.lbl_stats.pack(
            side="right"
        )

        controls = tk.Frame(
            self.root,
            bg="#1e293b",
            padx=14,
            pady=11,
            bd=1,
            relief="solid",
        )
        controls.pack(
            fill="x",
            padx=12,
            pady=(8, 5),
        )
        search_line = tk.Frame(controls, bg="#1e293b")
        search_line.pack(fill="x")
        scope_line = tk.Frame(controls, bg="#1e293b")
        scope_line.pack(fill="x", pady=(5, 0))
        actions = tk.Frame(controls, bg="#1e293b")
        actions.pack(fill="x", pady=(5, 0))

        tk.Label(
            search_line,
            text="SEARCH (MAX 18):",
            font=("Segoe UI", 11, "bold"),
            fg="#22d3ee",
            bg="#1e293b",
        ).pack(
            side="left",
            padx=(0, 7),
        )

        self.entry = tk.Entry(
            search_line,
            font=("Segoe UI", 13),
            width=24,
            bg="#0b1220",
            fg="#ffffff",
            insertbackground="#22d3ee",
            relief="flat",
            bd=2,
            validate="key",
            validatecommand=(self.root.register(
                lambda proposed: len(proposed) <= 18 and all(char.isprintable() for char in proposed)
            ), "%P"),
        )
        self.entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=5,
            ipady=5,
        )

        self.entry.bind(
            "<Return>",
            lambda _e: self.search_with_refresh(),
        )
        self.entry.bind("<KeyRelease>", self.on_typing)
        self.entry.bind("<<Paste>>", lambda _e: self.root.after_idle(self.on_typing))

        for label, scope in (("Data Inside Files", "data"),
                             ("File Name / Path", "file")):
            tk.Radiobutton(scope_line, text=label, variable=self.search_scope,
                           value=scope, command=self.change_search_scope,
                           bg="#1e293b", fg="#e2e8f0", selectcolor="#334155",
                           activebackground="#1e293b", activeforeground="#22d3ee",
                           font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 18))

        tk.Button(
            search_line,
            text="🔎 Search",
            font=("Segoe UI", 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            padx=14,
            pady=5,
            relief="flat",
            command=self.search_with_refresh,
        ).pack(
            side="left",
            padx=4,
        )

        tk.Button(
            search_line,
            text="✖ Clear",
            font=("Segoe UI", 9, "bold"),
            bg="#475569",
            fg="#ffffff",
            padx=11,
            pady=5,
            relief="flat",
            command=self.clear_search,
        ).pack(
            side="left",
            padx=4,
        )

        tk.Button(
            actions,
            text="📂 Select Data Folder",
            font=("Segoe UI", 9, "bold"),
            bg="#0f766e",
            fg="#ffffff",
            padx=12,
            pady=5,
            relief="flat",
            command=self.choose_folder,
        ).pack(
            side="left",
            padx=4,
        )

        tk.Button(
            actions, text="Connect to PC", command=self.connect_pc,
            font=("Segoe UI", 9, "bold"), bg="#7c3aed", fg="#ffffff",
            padx=12, pady=5, relief="flat",
        ).pack(side="left", padx=4)

        tk.Button(
            actions,
            text="🔄 Rescan",
            font=("Segoe UI", 9, "bold"),
            bg="#10b981",
            fg="#ffffff",
            padx=12,
            pady=5,
            relief="flat",
            command=self.start_scan,
        ).pack(
            side="right",
            padx=4,
        )

        tk.Button(
            actions,
            text="📄 Export ALL",
            font=("Segoe UI", 9, "bold"),
            bg="#e11d48",
            fg="#ffffff",
            padx=13,
            pady=5,
            relief="flat",
            command=self.export_all,
        ).pack(
            side="right",
            padx=4,
        )

        self.next_button = tk.Button(
            actions, text="Next ▶", command=lambda: self.change_page(1),
            font=("Segoe UI", 9, "bold"), bg="#2563eb", fg="#ffffff",
            relief="flat", state="disabled", padx=9, pady=5,
        )
        self.next_button.pack(side="right", padx=4)
        self.prev_button = tk.Button(
            actions, text="◀ Previous", command=lambda: self.change_page(-1),
            font=("Segoe UI", 9, "bold"), bg="#334155", fg="#ffffff",
            relief="flat", state="disabled", padx=9, pady=5,
        )
        self.prev_button.pack(side="right", padx=4)

        cloudbar = tk.Frame(self.root, bg="#16283b", padx=12, pady=5)
        cloudbar.pack(fill="x", padx=12, pady=(0, 5))
        cloud_actions = tk.Frame(cloudbar, bg="#16283b")
        cloud_actions.pack(fill="x")
        cloud_status = tk.Frame(cloudbar, bg="#16283b")
        cloud_status.pack(fill="x")
        tk.Button(
            cloud_actions, text="Connect Gmail + Drive", command=self.connect_google,
            bg="#2563eb", fg="#ffffff", relief="flat", padx=9,
        ).pack(side="left", padx=(0, 5))
        tk.Button(
            cloud_actions, text="Sync Google now", command=self.start_cloud_sync,
            bg="#0f766e", fg="#ffffff", relief="flat", padx=9,
        ).pack(side="left", padx=(0, 5))
        tk.Button(
            cloud_actions, text="Change Google account", command=self.change_google_account,
            bg="#334155", fg="#ffffff", relief="flat", padx=9,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            cloud_actions, text="OAuth settings", command=self.open_google_setup,
            bg="#334155", fg="#ffffff", relief="flat", padx=9,
        ).pack(side="left", padx=(0, 8))
        self.lbl_gmail = tk.Label(
            cloud_status, text="Gmail: not connected", bg="#16283b", fg="#fbbf24",
            font=("Segoe UI", 9, "bold"),
        )
        self.lbl_gmail.pack(side="left", padx=5)
        self.lbl_drive = tk.Label(
            cloud_status, text="Drive: not connected", bg="#16283b", fg="#fbbf24",
            font=("Segoe UI", 9, "bold"),
        )
        self.lbl_drive.pack(side="left", padx=5)

        info = tk.Frame(
            self.root,
            bg="#111827",
            padx=12,
            pady=8,
        )
        info.pack(
            fill="x",
            padx=12,
            pady=(0, 5),
        )

        self.lbl_status = tk.Label(
            info,
            text="Loading...",
            font=("Segoe UI", 9),
            fg="#cbd5e1",
            bg="#111827",
            anchor="w",
        )
        self.lbl_status.pack(
            fill="x"
        )

        self.lbl_folder = tk.Label(
            info,
            text=("PC CONNECTED: local drives + data folder" if self.pc_connected
                  else f"DATA FOLDER: {self.folder}"),
            font=("Segoe UI", 8),
            fg="#64748b",
            bg="#111827",
            anchor="w",
        )
        self.lbl_folder.pack(
            fill="x",
            pady=(3, 0),
        )

        self.progress = ttk.Progressbar(
            info,
            style="Fast.Horizontal.TProgressbar",
            mode="indeterminate",
        )
        self.progress.pack(
            fill="x",
            pady=(7, 0),
        )

        body = tk.Frame(
            self.root,
            bg="#080d16",
        )
        body.pack(
            fill="both",
            expand=True,
            padx=12,
            pady=(0, 12),
        )

        body_split = tk.PanedWindow(
            body, orient="horizontal", bg="#080d16", sashwidth=7,
            showhandle=True, relief="flat", opaqueresize=True,
        )
        body_split.pack(fill="both", expand=True)

        results = tk.Frame(
            body_split,
            bg="#111827",
            bd=1,
            relief="solid",
        )
        body_split.add(results, minsize=500, stretch="always")

        rh = tk.Frame(
            results,
            bg="#1f2937",
            padx=10,
            pady=7,
        )
        rh.pack(fill="x")

        self.lbl_results = tk.Label(
            rh,
            text="📋 MATCHING ROWS",
            font=("Segoe UI", 10, "bold"),
            fg="#f8fafc",
            bg="#1f2937",
        )
        self.lbl_results.pack(
            side="left"
        )

        tk.Label(
            rh,
            text="Double-click = complete row + all columns",
            font=("Segoe UI", 8),
            fg="#94a3b8",
            bg="#1f2937",
        ).pack(
            side="right"
        )

        tf = tk.Frame(
            results,
            bg="#111827",
        )
        tf.pack(
            fill="both",
            expand=True,
        )

        sy = ttk.Scrollbar(
            tf,
            orient="vertical",
        )
        sx = ttk.Scrollbar(
            tf,
            orient="horizontal",
        )

        self.tree = ttk.Treeview(
            tf,
            columns=(
                "file",
                "type",
                "location",
                "matched",
                "complete",
            ),
            show="headings",
            yscrollcommand=sy.set,
            xscrollcommand=sx.set,
        )

        sy.config(
            command=self.tree.yview
        )
        sx.config(
            command=self.tree.xview
        )

        sy.pack(
            side="right",
            fill="y",
        )
        sx.pack(
            side="bottom",
            fill="x",
        )
        self.tree.pack(
            fill="both",
            expand=True,
        )

        for cid, title, width in [
            ("file", "FILE", 240),
            ("type", "TYPE", 70),
            ("location", "SHEET / ROW", 220),
            ("matched", "🎯 MATCHED CELL(S)", 380),
            ("complete", "COMPLETE ROW — ALL COLUMNS", 900),
        ]:
            self.tree.heading(
                cid,
                text=title,
            )
            self.tree.column(
                cid,
                width=width,
                anchor="w",
            )

        self.tree.bind(
            "<Double-1>",
            self.show_detail,
        )
        self.tree.bind("<<TreeviewSelect>>", self.show_row_preview)

        preview = tk.Frame(results, bg="#111827", height=180)
        preview.pack(side="bottom", fill="x", padx=8, pady=(3, 8))
        preview.pack_propagate(False)
        def resize_preview(event):
            desired = max(100, min(205, event.height // 4))
            if abs(preview.winfo_height() - desired) > 5:
                preview.configure(height=desired)
        results.bind("<Configure>", resize_preview)
        self.preview_title = tk.Label(
            preview, text="COMPLETE ROW | select a matching result",
            font=("Segoe UI", 9, "bold"), fg="#22d3ee",
            bg="#111827", anchor="w",
        )
        self.preview_title.pack(fill="x", pady=(0, 4))
        preview_table = tk.Frame(preview, bg="#111827")
        preview_table.pack(fill="both", expand=True)
        preview_scroll = ttk.Scrollbar(preview_table, orient="vertical")
        self.preview_tree = ttk.Treeview(
            preview_table, columns=("column", "value", "match"),
            show="headings", yscrollcommand=preview_scroll.set,
        )
        preview_scroll.configure(command=self.preview_tree.yview)
        preview_scroll.pack(side="right", fill="y")
        self.preview_tree.pack(fill="both", expand=True)
        for key, label, width in [
            ("column", "COLUMN", 110), ("value", "COMPLETE CELL VALUE", 700),
            ("match", "MATCH", 90),
        ]:
            self.preview_tree.heading(key, text=label)
            self.preview_tree.column(key, width=width, anchor="w")

        files_panel = tk.Frame(
            body_split,
            bg="#111827",
            width=345,
            bd=1,
            relief="solid",
        )
        body_split.add(files_panel, minsize=270, stretch="always")

        right_tabs = ttk.Notebook(files_panel)
        right_tabs.pack(fill="both", expand=True)
        live_tab = tk.Frame(right_tabs, bg="#111827")
        side_split = tk.PanedWindow(live_tab, orient="vertical", bg="#111827",
                                    sashwidth=6, showhandle=True, relief="flat")
        side_split.pack(fill="both", expand=True)
        history_tab = tk.Frame(side_split, bg="#111827")
        checked_tab = tk.Frame(right_tabs, bg="#111827")
        activity_tab = tk.Frame(side_split, bg="#111827")
        files_tab = tk.Frame(right_tabs, bg="#111827")
        side_split.add(history_tab, minsize=125, stretch="always")
        side_split.add(activity_tab, minsize=125, stretch="always")
        right_tabs.add(live_tab, text="History + logs")
        right_tabs.add(checked_tab, text="Checked numbers")
        right_tabs.add(files_tab, text="Data files")

        tk.Label(history_tab, text="LAST SEARCHED NUMBERS",
                 bg="#111827", fg="#22d3ee",
                 font=("Segoe UI", 9, "bold"), pady=8).pack(fill="x")
        self.history_tree = ttk.Treeview(
            history_tab, columns=("when", "number", "matches", "speed", "first", "times"),
            show="headings",
        )
        for key, label, width in [
            ("when", "TIME", 95), ("number", "QUERY / MODE", 115),
            ("matches", "RESULTS", 65), ("speed", "MS", 45),
            ("first", "FIRST SEARCH", 128), ("times", "TIMES", 52),
        ]:
            self.history_tree.heading(key, text=label)
            self.history_tree.column(key, width=width, anchor="w")
        history_scroll = ttk.Scrollbar(
            history_tab, orient="vertical", command=self.history_tree.yview
        )
        history_xscroll = ttk.Scrollbar(
            history_tab, orient="horizontal", command=self.history_tree.xview)
        self.history_tree.configure(yscrollcommand=history_scroll.set,
                                    xscrollcommand=history_xscroll.set)
        history_scroll.pack(side="right", fill="y")
        history_xscroll.pack(side="bottom", fill="x")
        self.history_tree.pack(fill="both", expand=True)
        self.history_tree.bind("<Double-1>", self.search_history_item)

        checked_controls = tk.Frame(checked_tab, bg="#111827", pady=7)
        checked_controls.pack(fill="x")
        self.checked_filter = tk.Entry(checked_controls, width=16)
        self.checked_filter.pack(side="left", padx=5)
        tk.Button(checked_controls, text="Find number", command=self.refresh_checked_numbers).pack(side="left")
        tk.Button(checked_controls, text="All", command=self.clear_checked_filter).pack(side="left", padx=4)
        self.checked_count = tk.Label(checked_tab, text="", bg="#111827", fg="#22d3ee")
        self.checked_count.pack(fill="x")
        checked_table = tk.Frame(checked_tab, bg="#111827")
        checked_table.pack(fill="both", expand=True)
        checked_scroll = ttk.Scrollbar(checked_table, orient="vertical")
        self.checked_tree = ttk.Treeview(
            checked_table, columns=("number", "first", "last", "times", "matches"),
            show="headings", yscrollcommand=checked_scroll.set,
        )
        checked_scroll.config(command=self.checked_tree.yview)
        checked_scroll.pack(side="right", fill="y")
        for key, label, width in (
            ("number", "NUMBER", 105), ("first", "FIRST SEARCH", 140),
            ("last", "LAST SEARCH", 140), ("times", "TIMES", 52),
            ("matches", "ROWS", 70),
        ):
            self.checked_tree.heading(key, text=label)
            self.checked_tree.column(key, width=width, anchor="w")
        self.checked_tree.pack(fill="both", expand=True)
        self.checked_tree.bind("<Double-1>", self.search_checked_number)
        checked_pages = tk.Frame(checked_tab, bg="#111827")
        checked_pages.pack(fill="x")
        tk.Button(checked_pages, text="◀ Previous", command=lambda: self.change_checked_page(-1)).pack(side="left")
        tk.Button(checked_pages, text="Next ▶", command=lambda: self.change_checked_page(1)).pack(side="right")

        tk.Label(activity_tab, text="SCAN AND SEARCH ACTIVITY",
                 bg="#111827", fg="#22d3ee",
                 font=("Segoe UI", 9, "bold"), pady=8).pack(fill="x")
        log_scroll = ttk.Scrollbar(activity_tab, orient="vertical")
        self.log_text = tk.Text(
            activity_tab, wrap="word", bg="#0b1220", fg="#cbd5e1",
            font=("Consolas", 9), relief="flat", state="disabled",
            yscrollcommand=log_scroll.set,
        )
        log_scroll.configure(command=self.log_text.yview)
        log_scroll.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        fh = tk.Frame(
            files_tab,
            bg="#1f2937",
            padx=10,
            pady=7,
        )
        fh.pack(
            fill="x"
        )

        tk.Label(
            fh,
            text="📁 SELECTED FOLDER FILES",
            font=("Segoe UI", 9, "bold"),
            fg="#22d3ee",
            bg="#1f2937",
        ).pack(
            side="left"
        )

        self.lbl_file_count = tk.Label(
            fh,
            text="0",
            font=("Segoe UI", 9, "bold"),
            fg="#34d399",
            bg="#1f2937",
        )
        self.lbl_file_count.pack(
            side="right"
        )

        ff = tk.Frame(
            files_tab,
            bg="#111827",
        )
        ff.pack(
            fill="both",
            expand=True,
        )

        fsy = ttk.Scrollbar(
            ff,
            orient="vertical",
        )

        self.file_tree = ttk.Treeview(
            ff,
            columns=(
                "name",
                "size",
            ),
            show="headings",
            yscrollcommand=fsy.set,
        )

        fsy.config(
            command=self.file_tree.yview
        )

        fsy.pack(
            side="right",
            fill="y",
        )

        self.file_tree.pack(
            fill="both",
            expand=True,
        )

        self.file_tree.heading(
            "name",
            text="FILE"
        )
        self.file_tree.heading(
            "size",
            text="SIZE"
        )

        self.file_tree.column(
            "name",
            width=225,
            anchor="w",
        )
        self.file_tree.column(
            "size",
            width=75,
            anchor="center",
        )

    def add_activity(self, message, persist=True):
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{stamp}] {message}\n"
        self.log_text.config(state="normal")
        self.log_text.insert("end", line)
        if int(self.log_text.index("end-1c").split(".")[0]) > 500:
            self.log_text.delete("1.0", "2.0")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        if persist:
            try:
                with self.log_path.open("a", encoding="utf-8") as f:
                    f.write(line)
            except OSError:
                pass

    def load_activity(self):
        if self.log_path.exists():
            try:
                with self.log_path.open("r", encoding="utf-8") as f:
                    for line in deque(f, maxlen=100):
                        self.log_text.config(state="normal")
                        self.log_text.insert("end", line)
                        self.log_text.config(state="disabled")
            except OSError:
                pass
        self.add_activity("Software ready. Checking data folder.")

    def load_history(self):
        self.index.import_history(self.history_path)
        if not self.history_path.exists():
            return
        try:
            with self.history_path.open("r", encoding="utf-8") as f:
                recent = deque(f, maxlen=200)
            for line in reversed(recent):
                item = json.loads(line)
                record = self.index.previous_query(item["query"])
                self.history_tree.insert("", "end", values=(
                    item["time"], item["query"], item["matches"], item["ms"],
                    record[0] if record else "", record[2] if record else 1,
                ))
        except (OSError, ValueError, KeyError):
            self.add_activity("Saved search history could not be loaded.")

    def record_search(self, query, count, more, elapsed):
        if not query:
            return
        row_count = f"{count}+" if more else str(count)
        history_query = ("FILE:" + query if self.current_scope == "file" else query)
        record = self.index.record_query(history_query, row_count)
        item = {
            "time": datetime.now().strftime("%m-%d %H:%M"),
            "query": history_query, "matches": row_count,
            "ms": round(elapsed * 1000),
            "first": record[0], "times": record[2],
        }
        self.history_tree.insert("", 0, values=(
            item["time"], item["query"], item["matches"], item["ms"],
            item["first"], item["times"],
        ))
        children = self.history_tree.get_children()
        for old in children[200:]:
            self.history_tree.delete(old)
        try:
            with self.history_path.open("a", encoding="utf-8") as f:
                f.write(json.dumps(item, ensure_ascii=False) + "\n")
        except OSError:
            self.add_activity("Search history could not be saved.")
        self.refresh_checked_numbers()

    def search_history_item(self, _event=None):
        selected = self.history_tree.selection()
        if selected:
            query = self.history_tree.item(selected[0], "values")[1]
            self.search_scope.set("file" if query.startswith("FILE:") else "data")
            if query.startswith("FILE:"):
                query = query[5:]
            self.entry.delete(0, "end")
            self.entry.insert(0, query)
            self.search_with_refresh()

    def clear_checked_filter(self):
        self.checked_filter.delete(0, "end")
        self.checked_offset = 0
        self.refresh_checked_numbers()

    def change_checked_page(self, direction):
        self.checked_offset = max(0, self.checked_offset + direction * 300)
        self.refresh_checked_numbers()

    def refresh_checked_numbers(self):
        exact = self.checked_filter.get().strip()
        rows, total = self.index.list_checked_numbers(self.checked_offset, 300, exact)
        if self.checked_offset >= total and not exact:
            self.checked_offset = max(0, ((max(0, total - 1)) // 300) * 300)
            rows, total = self.index.list_checked_numbers(self.checked_offset, 300)
        for item in self.checked_tree.get_children():
            self.checked_tree.delete(item)
        for row in rows:
            self.checked_tree.insert("", "end", values=row)
        self.checked_count.config(
            text=(f"{total:,} checked numbers | showing "
                  f"{(0 if exact else self.checked_offset) + 1 if total else 0:,}–"
                  f"{(0 if exact else self.checked_offset) + len(rows):,}")
        )

    def search_checked_number(self, _event=None):
        selected = self.checked_tree.selection()
        if selected:
            query = self.checked_tree.item(selected[0], "values")[0]
            self.entry.delete(0, "end")
            self.entry.insert(0, query)
            self.search_with_refresh()

    # ---------------------------------------------------------
    # GOOGLE GMAIL + DRIVE
    # ---------------------------------------------------------
    def google_progress(self, line):
        self.add_activity(line)
        if self.cloud_syncing:
            self.lbl_stats.config(text=line[:105], fg="#38bdf8")

    def connect_google(self):
        if self.cloud_connected:
            self.lbl_status.config(text="Google already connected. Click Sync Google now.")
            return
        if self.cloud.client_available():
            self.connect_google_saved()
            return
        self.open_google_setup()

    def open_google_setup(self):
        """Guide first-time OAuth registration; selecting JSON starts login."""
        current = getattr(self, "oauth_setup_dialog", None)
        if current and current.winfo_exists():
            current.lift()
            current.focus_force()
            return
        dialog = tk.Toplevel(self.root)
        self.oauth_setup_dialog = dialog
        dialog.title("Google Gmail + Drive connection")
        dialog.geometry("590x390")
        dialog.minsize(510, 345)
        dialog.configure(bg="#111827")
        dialog.transient(self.root)
        tk.Label(dialog, text="CONNECT GMAIL + GOOGLE DRIVE",
                 font=("Segoe UI", 13, "bold"), fg="#22d3ee",
                 bg="#111827", pady=12).pack(fill="x")
        tk.Label(dialog, text=(
            "First time only: Google requires your registered Desktop OAuth client.\n"
            "1. In your Google Cloud project enable Gmail, Drive and Sheets APIs.\n"
            "2. Configure OAuth consent; add your Gmail as a test user if in Testing.\n"
            "3. Create a Desktop OAuth client and download its JSON.\n"
            "4. Select that JSON below; your browser opens for account approval.\n\n"
            "After setup, Connect Gmail + Drive opens browser automatically\n"
            "when sign-in is needed. Never enter your password in this app."
        ), justify="left", font=("Segoe UI", 10), fg="#e5e7eb",
            bg="#111827", padx=20, pady=9).pack(fill="x")
        tk.Button(dialog, text="1. Open Google Cloud setup",
                  command=lambda: webbrowser.open("https://console.cloud.google.com/auth/clients"),
                  bg="#475569", fg="white", relief="flat", pady=7).pack(fill="x", padx=20, pady=5)

        def select_client():
            selected = filedialog.askopenfilename(
                parent=dialog, title="Select downloaded Google Desktop OAuth JSON",
                filetypes=[("Google OAuth JSON", "*.json")],
            )
            if selected:
                dialog.destroy()
                self.connect_google_client(selected)

        tk.Button(dialog, text="2. Select Desktop JSON and sign in",
                  command=select_client, bg="#2563eb", fg="white",
                  relief="flat", pady=8).pack(fill="x", padx=20, pady=5)
        tk.Button(dialog, text="Close", command=dialog.destroy,
                  bg="#334155", fg="white", relief="flat").pack(pady=5)

    def connect_google_client(self, client_json):
        if self.cloud_syncing or self.google_connecting:
            self.add_activity("Wait for the current Google connection or sync to finish.")
            return
        self.google_connecting = True
        self.lbl_gmail.config(text="Gmail: opening browser...", fg="#fbbf24")
        self.lbl_drive.config(text="Drive: opening browser...", fg="#fbbf24")
        threading.Thread(target=self.connect_google_worker,
                         args=(client_json,), daemon=True).start()

    def change_google_account(self):
        if self.cloud_syncing or self.google_connecting:
            self.add_activity("Wait for the current Google connection or sync to finish.")
            return
        if not self.cloud.client_available():
            self.open_google_setup()
            return
        self.cloud.token_path.unlink(missing_ok=True)
        self.cloud_connected = False
        self.connect_google_saved()

    def connect_google_saved(self):
        if self.google_connecting:
            return
        if not self.cloud.client_available():
            self.open_google_setup()
            return
        self.lbl_gmail.config(text="Gmail: verifying sign-in...", fg="#fbbf24")
        self.lbl_drive.config(text="Drive: verifying sign-in...", fg="#fbbf24")
        self.google_connecting = True
        threading.Thread(target=self.connect_google_worker,
                         args=(None,), daemon=True).start()

    def connect_google_worker(self, client_json):
        try:
            email = self.cloud.connect(client_json)
            self.root.after(0, self.google_connected, email)
        except Exception as exc:
            self.root.after(0, self.google_failed, str(exc))

    def google_failed(self, detail):
        self.google_connecting = False
        self.cloud_connected = False
        self.lbl_gmail.config(text="Gmail: connection failed", fg="#f87171")
        self.lbl_drive.config(text="Drive: connection failed", fg="#f87171")
        self.add_activity(f"Google connection failed: {detail}")
        self.lbl_status.config(text=f"Google connection failed: {detail}", fg="#f87171")
        messagebox.showerror("Google sign-in", detail)

    def google_connected(self, email):
        self.google_connecting = False
        self.cloud_connected = True
        self.cloud.data_folder = self.folder
        self.lbl_gmail.config(text=f"Gmail: connected ({email})", fg="#34d399")
        self.lbl_drive.config(text=f"Drive: connected ({email})", fg="#34d399")
        self.add_activity(f"Google account connected successfully: {email}")
        self.start_cloud_sync()

    def start_cloud_sync(self):
        if not self.cloud_connected:
            self.lbl_status.config(text="Connect Gmail + Drive first.", fg="#fbbf24")
            return
        if self.cloud_syncing:
            return
        if self.scanning:
            self.cloud_waiting = True
            self.add_activity("Google sync queued until the current folder scan finishes.")
            return
        self.cloud_waiting = False
        self.cloud_syncing = True
        if self.current_query:
            self.pending_query = self.current_query
        self.cloud.data_folder = self.folder
        self.lbl_status.config(text="Gmail and Drive syncing; indexed rows remain searchable.", fg="#fbbf24")
        self.progress["mode"] = "indeterminate"
        self.progress.start(12)
        self.lbl_stats.config(text="Gmail and Drive loading...", fg="#38bdf8")
        self.add_activity("Checking all Gmail messages and Drive files; no file-count cutoff.")
        threading.Thread(target=self.cloud_sync_worker, daemon=True).start()

    def cloud_sync_worker(self):
        try:
            result = self.cloud.sync()
        except Exception as exc:
            result = {"counts": {}, "errors": [str(exc)]}
        self.root.after(0, self.cloud_sync_done, result)

    def cloud_sync_done(self, result):
        self.cloud_syncing = False
        self.progress.stop()
        self.last_cloud_failed = len(result["errors"])
        self.add_activity(
            f"Google sync finished: Gmail {result['counts'].get('gmail', 0):,} messages, "
            f"Drive {result['counts'].get('drive', 0):,} files; "
            f"{len(result['errors']):,} errors."
        )
        if result["errors"]:
            messagebox.showwarning(
                "Google files need attention",
                f"{len(result['errors'])} item(s) could not be mirrored. "
                "See Activity logs; these items were not searched.\n\n"
                + "\n".join(result["errors"][:5]),
            )
        if self.scanning:
            self.sync_needs_scan = True
        else:
            self.sync_needs_scan = False
            if self.pending_pc_connect is True:
                self.pending_pc_connect = False
                self.connect_pc()
            else:
                self.start_scan()

    # ---------------------------------------------------------
    # FOLDER
    # ---------------------------------------------------------
    def choose_folder(self):
        if self.scanning or self.cloud_syncing:
            messagebox.showinfo("Loading in progress", "Wait for the current scan or Google sync to finish.")
            return
        selected = filedialog.askdirectory(
            initialdir=str(
                self.folder
            ),
            title=(
                "Select the folder containing your data files"
            ),
        )

        if not selected:
            return

        self.folder = Path(
            selected
        )
        self.pc_connected = False
        self.pc_state_path.write_text('{"connected": false}', encoding="utf-8")
        self.cloud.data_folder = self.folder
        self.folder_snapshot = None
        self.search_token += 1

        self.lbl_folder.config(
            text=f"DATA FOLDER: {self.folder}"
        )

        self.current_query = ""
        self.current_results = []
        self.page_starts = [0]
        self.page_index = 0
        self.has_more = False

        self.clear_results()

        self.refresh_file_panel()

        safe_name = re.sub(
            r"[^A-Za-z0-9]+",
            "_",
            self.folder.name,
        ).strip("_") or "folder"
        folder_key = hashlib.sha256(
            str(self.folder.resolve()).encode("utf-8")
        ).hexdigest()[:12]

        try:
            self.index.close()
        except Exception:
            pass

        self.index = FastIndex(
            self.base_dir
            / f"hacker_mt181_{safe_name}_{folder_key}.db"
        )
        self.index.import_history(self.history_path)
        self.checked_offset = 0
        self.refresh_checked_numbers()

        self.start_scan()

    def connect_pc(self):
        if os.name != "nt":
            self.lbl_status.config(text="Connect to PC requires Windows.", fg="#fbbf24")
            return
        if self.scanning or self.cloud_syncing or self.exporting:
            self.pending_pc_connect = True
            if self.scanning:
                self.cancel_scan.set()
            self.add_activity("Connect to PC queued until current work finishes.")
            return
        drives = local_pc_drives()
        if not drives:
            messagebox.showerror("Connect to PC", "No local disk drives were detected.")
            return
        if not self.pc_connected:
            self.search_token += 1
            self.clear_results()
            self.index.close()
            self.index = FastIndex(self.base_dir / "hacker_mt181_pc_index.db")
            self.index.import_history(self.history_path)
            self.pc_connected = True
            self.pc_state_path.write_text('{"connected": true}', encoding="utf-8")
            self.checked_offset = 0
            self.refresh_checked_numbers()
        self.folder_snapshot = None
        self.lbl_folder.config(text="PC CONNECTED: " + ", ".join(str(d) for d in drives)
                               + " | also scanning data folder")
        self.add_activity("PC connected: " + ", ".join(str(d) for d in drives)
                          + ". Reading supported documents; data stays on this PC.")
        self.start_scan()

    def refresh_file_panel(self):
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        if self.pc_connected:
            with self.index.lock:
                files = self.index.conn.execute(
                    "SELECT path,file_size FROM files WHERE indexed_ok IN (1,2) "
                    "ORDER BY id DESC LIMIT 150"
                ).fetchall()
                total = self.index.conn.execute(
                    "SELECT COUNT(*) FROM files WHERE indexed_ok IN (1,2)"
                ).fetchone()[0]
            for path, size in files:
                self.file_tree.insert("", "end", values=(path, f"{size/1024:.0f} KB"))
            self.lbl_file_count.config(text=f"{total:,} indexed; showing latest 150")
            return

        files = files_under(
            self.folder
        )

        for p in files:
            try:
                n = p.stat().st_size

                text = (
                    f"{n/1024/1024:.1f} MB"
                    if n >= 1024 * 1024
                    else f"{n/1024:.0f} KB"
                )
            except Exception:
                text = "-"

            try:
                display = str(
                    p.relative_to(
                        self.folder
                    )
                )
            except Exception:
                display = p.name

            self.file_tree.insert(
                "",
                "end",
                values=(
                    display,
                    text,
                ),
            )

        self.lbl_file_count.config(
            text=f"{len(files):,}"
        )

    # ---------------------------------------------------------
    # SCAN
    # ---------------------------------------------------------
    def folder_signature(self):
        signature = []
        for path in files_under(self.folder):
            try:
                st = path.stat()
                signature.append((str(path.resolve()), st.st_mtime_ns, st.st_size))
            except OSError:
                pass
        return tuple(signature)

    def check_folder_changes(self):
        try:
            if not self.pc_connected and not self.scanning and not self.cloud_syncing and self.folder_snapshot is not None:
                if self.folder_signature() != self.folder_snapshot:
                    self.start_scan()
        finally:
            self.root.after(15000, self.check_folder_changes)

    def start_scan(self):
        if self.cloud_syncing:
            self.sync_needs_scan = True
            return
        if self.scanning or self.exporting:
            return

        self.refresh_file_panel()
        # Keep the pre-scan snapshot. If a file changes while the scan is
        # running, the next periodic check will detect it and scan again.
        self.folder_snapshot = None if self.pc_connected else self.folder_signature()

        self.scanning = True
        self.last_scan_failed = 0
        if self.current_query:
            self.pending_query = self.current_query
        self.add_activity("Scanning local PC drives" if self.pc_connected
                          else f"Scanning folder: {self.folder}")
        self.cancel_scan.clear()
        self.scan_started_at = time.perf_counter()
        self.completed_rows = 0
        self._pc_pending_completed_rows = 0
        self._pc_last_update = 0.0

        self.progress.stop()
        self.progress["mode"] = "indeterminate" if self.pc_connected else "determinate"
        self.progress["value"] = 0
        if self.pc_connected:
            self.progress.start(12)

        self.lbl_status.config(
            text=(
                "Discovering all supported files on local drives; indexed rows are searchable now."
                if self.pc_connected else
                "Loading every row from the data folder; indexed rows are searchable now."
            ),
            fg="#22d3ee",
        )

        self.lbl_stats.config(
            text="Scanning...",
            fg="#38bdf8",
        )

        threading.Thread(
            target=self.scan_worker,
            daemon=True,
        ).start()

    def scan_worker(self):
        started = time.perf_counter()
        try:
            roots = [*local_pc_drives(), self.folder] if self.pc_connected else self.folder
            if self.pc_connected and len(roots) == 1:
                raise RuntimeError("No local drives available; check connected disks.")
            if self.pc_connected:
                previous_data = self.base_dir.parent / "HackerMT181" / "data"
                if previous_data.is_dir() and previous_data.resolve() != self.folder.resolve():
                    roots.append(previous_data)
                    self.root.after(0, self.add_activity,
                                    f"Earlier installation data folder included: {previous_data}")
            result = self.index.scan(
                roots,
                progress=self.scan_callback,
                cancel=self.cancel_scan,
                pc_mode=self.pc_connected,
                exclude_dirs=(self.base_dir,) if self.pc_connected else (),
            )
        except Exception as exc:
            self.root.after(0, self.scan_crashed, str(exc))
            return

        elapsed = (
            time.perf_counter()
            - started
        )

        self.root.after(
            0,
            self.scan_done,
            result,
            elapsed,
        )

    def scan_callback(
        self,
        kind,
        number,
        total,
        path,
        rows,
        detail="",
    ):
        prior_rows = 0
        if self.pc_connected and kind in {"start", "reading", "done", "cached", "copied"}:
            now = time.monotonic()
            if now - self._pc_last_update < 0.35 and number != total:
                if kind in {"done", "cached", "copied"}:
                    self._pc_pending_completed_rows += rows
                return
            self._pc_last_update = now
        if self.pc_connected and kind not in {"discovery", "directory_error"}:
            prior_rows = self._pc_pending_completed_rows
            self._pc_pending_completed_rows = 0
        self.root.after(
            0,
            self.scan_ui,
            kind,
            number,
            total,
            str(path),
            rows,
            detail,
            prior_rows,
        )

    def scan_ui(
        self,
        kind,
        number,
        total,
        path,
        rows,
        detail="",
        prior_rows=0,
    ):
        if kind == "discovery":
            self.lbl_status.config(
                text=f"Finding files: {number:,} folders checked; {rows:,} supported files found...",
                fg="#22d3ee")
            self.lbl_stats.config(text=f"PC FILE DISCOVERY | {rows:,} files", fg="#38bdf8")
            return
        if kind == "directory_error":
            self.add_activity(f"UNREADABLE FOLDER: {path}: {detail}")
            return
        if self.progress["mode"] == "indeterminate":
            self.progress.stop()
            self.progress["mode"] = "determinate"
        self.completed_rows += prior_rows
        name = Path(
            path
        ).name
        completed = kind in {"cached", "done", "copied", "error"}
        if completed:
            self.completed_rows += rows
        current_rows = self.completed_rows if completed else self.completed_rows + rows
        percent = (number if completed else number - 1) * 100 / max(1, total)
        self.progress["value"] = percent
        elapsed = max(0.001, time.perf_counter() - self.scan_started_at)

        if kind == "start":
            message = (
                f"📖 [{number}/{total}] "
                f"Reading {name}..."
            )

        elif kind == "reading":
            message = (
                f"⚡ [{number}/{total}] "
                f"{name} | {rows:,} rows loaded"
            )

        elif kind == "done":
            message = (
                f"✅ [{number}/{total}] "
                f"{name} | {rows:,} rows indexed"
            )

        elif kind == "cached":
            message = (
                f"⚡ [{number}/{total}] "
                f"{name} | reused {rows:,} rows"
            )

        elif kind == "copied":
            message = f"✅ [{number}/{total}] {name} | {rows:,} rows reused from identical file"

        else:
            message = (
                f"⚠ [{number}/{total}] "
                f"{name} failed: {detail}"
            )

        self.lbl_status.config(
            text=message,
            fg=(
                "#fbbf24"
                if kind == "error"
                else "#22d3ee"
            ),
        )

        if kind == "error" or (kind in {"start", "done", "copied"}
                               and (not self.pc_connected or number == 1 or number == total
                                    or number % 200 == 0)):
            self.add_activity(message)

        self.lbl_stats.config(
            text=(
                f"FILES {number}/{total} | {percent:.0f}% complete | "
                f"{current_rows:,} rows read | {current_rows/elapsed:,.0f} rows/s"
            ),
            fg="#38bdf8",
        )
        if self.current_query and kind in {"reading", "done", "cached", "copied"}:
            now = time.monotonic()
            if now - self.last_live_refresh > 3:
                self.last_live_refresh = now
                self.start_search(self.search_mode)

    def scan_crashed(self, detail):
        self.scanning = False
        self.last_scan_failed = 1
        self.progress.stop()
        self.lbl_status.config(text=f"SCAN ERROR: {detail}", fg="#f87171")
        self.lbl_stats.config(text="SCAN ERROR | See Activity log", fg="#f87171")
        self.add_activity(f"SCAN ERROR: {detail}")
        messagebox.showerror("Scan could not finish", detail)
        if self.pending_pc_connect is True:
            self.pending_pc_connect = False
            self.root.after(0, self.connect_pc)

    def scan_done(
        self,
        result,
        elapsed,
    ):
        self.scanning = False
        if self.pending_pc_connect is True:
            self.pending_pc_connect = False
            self.progress.stop()
            self.root.after(0, self.connect_pc)
            return
        self.last_scan_failed = result["failed"]
        self.add_activity(
            f"Scan finished: {result['ok_files']}/{result['files']} files; "
            f"{result['rows']:,} rows; {result['failed']} failed; {elapsed:.1f}s"
        )

        self.progress.stop()
        self.progress["mode"] = "determinate"
        self.progress["value"] = 100

        self.refresh_file_panel()

        self.lbl_stats.config(
            text=(
                f"READY | {result['ok_files']:,}/"
                f"{result['files']:,} FILES | "
                f"{result['rows']:,} ROWS | "
                f"{elapsed:.1f}s"
            ),
            fg="#34d399"
            if not result["failed"]
            else "#fbbf24",
        )

        text = (
            f"✅ Scan complete in {elapsed:.1f}s — "
            f"{result['ok_files']:,} files / "
            f"{result['rows']:,} rows indexed."
        )

        if result["failed"]:
            text += (
                f" ⚠ {result['failed']} file(s) failed."
            )

        self.lbl_status.config(
            text=text,
            fg="#34d399"
            if not result["failed"]
            else "#fbbf24",
        )
        for error in result["errors"]:
            self.add_activity(f"FILE ERROR: {error}")
        if result["errors"]:
            messagebox.showwarning(
                "Some files could not be read",
                "These files were not searched:\n\n"
                + "\n".join(result["errors"][:20]),
            )

        if self.cloud_waiting:
            self.start_cloud_sync()
            return
        if self.sync_needs_scan:
            self.sync_needs_scan = False
            self.start_scan()
            return

        if self.pending_query or self.current_query:
            self.pending_query = ""
            self.start_search(self.search_mode)

    # ---------------------------------------------------------
    # SEARCH
    # ---------------------------------------------------------
    def change_search_scope(self):
        self.search_token += 1
        if self.entry.get().strip():
            self.start_search("preview")
        else:
            self.lbl_results.config(text=("📁 MATCHING FILE NAMES / PATHS"
                                          if self.search_scope.get() == "file"
                                          else "📋 MATCHING CONTENT ROWS"))

    def search_with_refresh(self):
        query = self.entry.get().strip()
        if not self.pc_connected and not self.scanning and self.folder_snapshot != self.folder_signature():
            self.start_scan()
        if self.search_after:
            self.root.after_cancel(self.search_after)
            self.search_after = None
        self.start_search("exact")

    def on_typing(self, event=None):
        if event and event.keysym in {
            "Return", "Tab", "Escape",
            "Up",
            "Down",
            "Left",
            "Right",
            "Shift_L",
            "Shift_R",
            "Control_L",
            "Control_R",
            "Alt_L",
            "Alt_R",
            "Caps_Lock",
        }:
            return

        self.search_token += 1  # Discard responses for an older entry value.

        if self.search_after:
            try:
                self.root.after_cancel(
                    self.search_after
                )
            except Exception:
                pass

        self.search_after = self.root.after(
            180,
            lambda: self.start_search("preview"),
        )

    def start_search(self, mode="exact"):
        self.search_after = None
        query = self.entry.get().strip()

        if not query:
            self.clear_search()
            return
        if len(query) > 18:
            self.search_token += 1
            self.pending_query = ""
            self.current_query = ""
            self.clear_results()
            self.lbl_results.config(text="Search text is limited to 18 characters")
            self.lbl_status.config(text="Enter up to 18 characters.", fg="#fbbf24")
            return

        self.search_mode = mode
        self.current_scope = self.search_scope.get()
        self.page_starts = [(len(query), 0)] if mode == "preview" else [0]
        self.page_index = 0
        self.next_cursor = None
        self.search_token += 1
        token = self.search_token

        self.current_query = query

        self.clear_results()

        started = time.perf_counter()

        self.pending_query = query if self.scanning or self.cloud_syncing else ""
        self.previous_search = self.index.previous_query(
            "FILE:" + query if self.current_scope == "file" else query)
        if self.previous_search and mode == "exact":
            self.add_activity(
                f"{query} was previously searched {self.previous_search[2]} time(s); "
                f"first {self.previous_search[0]}, last {self.previous_search[1]}"
            )
        if mode == "exact":
            self.add_activity(f"Searching {'file names and paths' if self.current_scope == 'file' else 'content inside files'}: {query}")
        else:
            self.lbl_status.config(text=f"Live search: {query} | checking indexed rows...", fg="#38bdf8")

        threading.Thread(
            target=self.search_worker,
            args=(
                query,
                token,
                started,
                self.page_index,
                self.page_starts[self.page_index],
                mode,
            ),
            daemon=True,
        ).start()

    def change_page(self, direction):
        if not self.current_query:
            return
        if direction > 0:
            if not self.has_more or not self.current_results:
                return
            next_start = (self.next_cursor if self.search_mode == "preview"
                          else self.current_results[-1]["id"])
            if len(self.page_starts) <= self.page_index + 1:
                self.page_starts.append(next_start)
            self.page_index += 1
        elif self.page_index > 0:
            self.page_index -= 1
        else:
            return
        self.search_token += 1
        token = self.search_token
        self.clear_results()
        threading.Thread(
            target=self.search_worker,
            args=(self.current_query, token, time.perf_counter(),
                  self.page_index, self.page_starts[self.page_index], self.search_mode),
            daemon=True,
        ).start()

    def search_worker(
        self,
        query,
        token,
        started,
        page_index,
        after_id,
        mode="exact",
    ):
        try:
            if getattr(self, "current_scope", "data") == "file":
                cursor = after_id[1] if isinstance(after_id, tuple) else after_id
                results, has_more = self.index.search_files_page(query, cursor)
                next_cursor = (len(query), results[-1]["id"]) if results else None
            elif mode == "preview" and not self.scanning and not self.cloud_syncing:
                def show_exact(page, more):
                    if (page_index == 0 and len(page) < PAGE_SIZE and len(query) > 4
                            and query.isascii() and query.isdigit()):
                        self.root.after(0, self.search_done, query, token, page,
                                        time.perf_counter() - started, page_index,
                                        False, None, mode, True)
                results, has_more, next_cursor = self.index.search_ranked_page(
                    query, after_id, cancelled=lambda: token != self.search_token,
                    exact_ready=show_exact,
                )
            else:
                cursor = after_id[1] if isinstance(after_id, tuple) else after_id
                results, has_more = self.index.search_page(query, cursor)
                for row in results:
                    row["rank"] = len(query)
                next_cursor = (len(query), results[-1]["id"]) if results else None
        except (InterruptedError, sqlite3.OperationalError) as exc:
            if token != self.search_token:
                return
            self.root.after(0, self.search_crashed, token, query, str(exc))
            return
        except Exception as exc:
            self.root.after(0, self.search_crashed, token, query, str(exc))
            return

        elapsed = (
            time.perf_counter()
            - started
        )

        self.root.after(
            0,
            self.search_done,
            query,
            token,
            results,
            elapsed,
            page_index,
            has_more,
            next_cursor,
            mode,
        )

    def search_crashed(self, token, query, detail):
        if token != self.search_token:
            return
        self.lbl_status.config(text=f"SEARCH ERROR: {detail}", fg="#f87171")
        self.add_activity(f"SEARCH ERROR for {query}: {detail}")
        messagebox.showerror("Search could not finish", detail)

    def search_done(
        self,
        query,
        token,
        results,
        elapsed,
        page_index,
        has_more,
        next_cursor=None,
        mode="exact",
        incomplete=False,
    ):
        if token != self.search_token:
            return

        self.clear_results()
        self.current_results = results
        self.has_more = has_more
        self.next_cursor = next_cursor
        if page_index == 0 and mode == "exact" and not self.scanning and not self.cloud_syncing:
            self.record_search(query, len(results), has_more, elapsed)
            self.add_activity(
                f"Search {query}: {len(results)} matching {'files' if self.current_scope == 'file' else 'rows'} on first page"
                f"{' (more pages)' if has_more else ''}; {elapsed*1000:.0f} ms"
            )
        self.prev_button.config(state="normal" if page_index > 0 else "disabled")
        self.next_button.config(state="normal" if has_more and not incomplete else "disabled")

        total = len(
            results
        )

        shown = total

        for result in results:
            path = Path(
                result["path"]
            )
            if "__Google_Cloud__" in path.parts:
                anchor = path.parts.index("__Google_Cloud__")
                display_name = " / ".join(path.parts[anchor + 2:])
            else:
                display_name = path.name

            matched = " | ".join(
                f"[{c}]: {v}"
                for c, v
                in result["matched"]
            )
            if self.current_scope == "file":
                matched = "FILE MATCH | " + matched
            elif result.get("rank", len(query)) < len(query):
                matched = f"NEAR {result['rank']}/{len(query)} digits | " + matched
            elif mode == "preview" and len(query) >= 4 and query.isascii() and query.isdigit():
                matched = f"EXACT {len(query)}/{len(query)} digits | " + matched

            self.tree.insert(
                "",
                "end",
                values=(
                    display_name,
                    path.suffix.replace(
                        ".",
                        "",
                    ).upper(),
                    (result["path"] if self.current_scope == "file" else
                     f"Sheet: {result['sheet']} | Row: {result['row']}"),
                    matched,
                    result["full"],
                ),
            )

        if results:
            first = self.tree.get_children()[0]
            self.tree.selection_set(first)
            self.show_row_preview()

        if total:
            suffix = f" | Page {page_index + 1}, {shown:,} rows"
            if has_more:
                suffix += " | More matches: Next ▶"

            self.lbl_results.config(
                text=(
                    f"{'📁 FILE NAMES' if self.current_scope == 'file' else '📋 DATA INSIDE FILES'} | "
                    f"'{query}' | page {page_index + 1}: "
                    f"{shown:,} matching {'files' if self.current_scope == 'file' else 'rows'}" +
                    (" | near matches loading..." if incomplete else
                     " | more: Next ▶" if has_more else "")
                ),
            )

            if self.scanning:
                status = (
                    f"⚡ LIVE SEARCH {elapsed*1000:.1f} ms | "
                    f"{total:,} matches in rows loaded so far. "
                    "Full final result will refresh automatically after scan."
                )
                color = "#fbbf24"
            elif mode == "preview":
                status = (f"⚡ LIVE {elapsed*1000:.0f} ms | "
                          + ("matching file names and paths. " if self.current_scope == 'file'
                             else "matching content; " +
                                  ("numeric near matches down to 4 digits. "
                                   if query.isascii() and query.isdigit() else "exact text matches. "))
                          + "Click Search for exact results only.")
                if incomplete:
                    status += " Near matches still loading..."
                color = "#38bdf8"
            else:
                status = (
                    f"✅ FINAL SEARCH {elapsed*1000:.1f} ms | "
                    + ("matching file names and paths." if self.current_scope == 'file'
                       else "ANY cell match → COMPLETE row.")
                )
                if self.previous_search:
                    status += f" Previously searched: {self.previous_search[1]} ({self.previous_search[2]} times)."
                if self.last_scan_failed or self.last_cloud_failed:
                    status = (f"⚠ {self.last_scan_failed} local/parsed files and "
                              f"{self.last_cloud_failed} Google items failed. Results may be incomplete.")
                color = "#fbbf24" if self.last_scan_failed or self.last_cloud_failed else "#34d399"

            self.lbl_status.config(
                text=status + suffix,
                fg=color,
            )

        else:
            self.lbl_results.config(
                text=(
                    f"{'📁 FILE NAMES' if self.current_scope == 'file' else '📋 DATA INSIDE FILES'} | 0 MATCHES | "
                    f"'{query}'"
                ),
            )

            if self.scanning:
                status = (
                    f"🔎 LIVE SEARCH: '{query}' not found in rows loaded so far. "
                    "The full folder will be checked again automatically after scan."
                )
                color = "#fbbf24"
            elif incomplete:
                status = f"🔎 Exact '{query}' not found; checking nearby numbers..."
                color = "#38bdf8"
            else:
                status = (f"❌ {'LIVE' if mode == 'preview' else 'FINAL'}: "
                          f"'{query}' not found in the indexed files.")
                if self.previous_search:
                    status += f" Previously searched: {self.previous_search[1]}."
                if self.last_scan_failed or self.last_cloud_failed:
                    status += (f" WARNING: {self.last_scan_failed} parsed files, "
                               f"{self.last_cloud_failed} Google items could not be checked.")
                color = "#fbbf24" if self.last_scan_failed or self.last_cloud_failed else "#f87171"

            self.lbl_status.config(
                text=status,
                fg=color,
            )

    def clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if hasattr(self, "preview_tree"):
            for item in self.preview_tree.get_children():
                self.preview_tree.delete(item)
            self.preview_title.config(text="COMPLETE ROW | select a matching result")

    def show_row_preview(self, _event=None):
        selected = self.tree.selection()
        if not selected or not self.current_results:
            return
        idx = self.tree.index(selected[0])
        if idx >= len(self.current_results):
            return
        row = self.current_results[idx]
        self.preview_title.config(
            text=f"COMPLETE ROW | {row['path']} | {row['sheet']} | row {row['row']}"
        )
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        matched = {col for col, _value in row["matched"]}
        for col, value in row["cells"]:
            self.preview_tree.insert(
                "", "end", values=(col, value, "YES" if col in matched else "")
            )

    def clear_search(self):
        if self.search_after:
            try:
                self.root.after_cancel(self.search_after)
            except Exception:
                pass
            self.search_after = None
        self.search_token += 1
        self.pending_query = ""

        self.current_query = ""
        self.current_results = []
        self.page_starts = [0]
        self.page_index = 0
        self.has_more = False
        self.next_button.config(state="disabled")
        self.prev_button.config(state="disabled")

        self.entry.delete(
            0,
            "end",
        )

        self.clear_results()

        self.lbl_results.config(
            text="📋 MATCHING ROWS"
        )

        self.lbl_status.config(
            text="Search cleared.",
            fg="#94a3b8",
        )

        self.entry.focus_set()

    # ---------------------------------------------------------
    # DETAIL
    # ---------------------------------------------------------
    def show_detail(self, _event=None):
        selected = self.tree.selection()

        if not selected:
            return

        idx = self.tree.index(
            selected[0]
        )

        if idx >= len(
            self.current_results
        ):
            return

        r = self.current_results[
            idx
        ]

        win = tk.Toplevel(
            self.root
        )
        win.title(
            "Complete Row | HACKER M-T181"
        )
        win.geometry(
            "1080x730"
        )
        win.configure(
            bg="#0b0f19"
        )

        tk.Label(
            win,
            text=(
                f"📄 {Path(r['path']).name}"
            ),
            font=("Segoe UI", 13, "bold"),
            fg="#22d3ee",
            bg="#0b0f19",
            pady=10,
        ).pack(
            fill="x"
        )

        tk.Label(
            win,
            text=(
                f"📍 Sheet: {r['sheet']} | "
                f"Row: {r['row']}"
            ),
            font=("Segoe UI", 10, "bold"),
            fg="#34d399",
            bg="#0b0f19",
        ).pack(
            anchor="w",
            padx=15,
        )

        tk.Label(
            win,
            text=(
                "🎯 MATCHED CELL(S): "
                + " | ".join(
                    f"[{c}]: {v}"
                    for c, v
                    in r["matched"]
                )
            ),
            font=("Segoe UI", 10, "bold"),
            fg="#fbbf24",
            bg="#0b0f19",
            wraplength=1000,
            justify="left",
        ).pack(
            anchor="w",
            padx=15,
            pady=7,
        )

        frame = tk.Frame(
            win,
            bg="#111827",
        )
        frame.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=10,
        )

        sy = ttk.Scrollbar(
            frame,
            orient="vertical",
        )
        sx = ttk.Scrollbar(
            frame,
            orient="horizontal",
        )

        tv = ttk.Treeview(
            frame,
            columns=(
                "column",
                "value",
                "matched",
            ),
            show="headings",
            yscrollcommand=sy.set,
            xscrollcommand=sx.set,
        )

        sy.config(
            command=tv.yview
        )
        sx.config(
            command=tv.xview
        )

        sy.pack(
            side="right",
            fill="y",
        )
        sx.pack(
            side="bottom",
            fill="x",
        )
        tv.pack(
            fill="both",
            expand=True,
        )

        tv.heading(
            "column",
            text="COLUMN"
        )
        tv.heading(
            "value",
            text="CELL VALUE"
        )
        tv.heading(
            "matched",
            text="MATCH"
        )

        tv.column(
            "column",
            width=100,
            anchor="center",
        )
        tv.column(
            "value",
            width=820,
            anchor="w",
        )
        tv.column(
            "matched",
            width=80,
            anchor="center",
        )

        matched_cols = {
            c
            for c, _
            in r["matched"]
        }

        for col, value in r["cells"]:
            tv.insert(
                "",
                "end",
                values=(
                    col,
                    value,
                    "YES"
                    if col in matched_cols
                    else "",
                ),
            )

    # ---------------------------------------------------------
    # EXPORT
    # ---------------------------------------------------------
    def export_all(self):
        if self.scanning:
            messagebox.showinfo("Scan in progress", "Wait until all files finish loading, then export.")
            return
        if self.exporting:
            return
        if not self.current_query:
            messagebox.showwarning(
                "No Search",
                "Pehle search karein."
            )
            return

        if not self.current_results:
            messagebox.showwarning(
                "No Results",
                "Koi matching result nahi."
            )
            return

        save = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[
                ("CSV", "*.csv")
            ],
            initialfile=(
                "SearchReport_"
                + re.sub(
                    r"[^A-Za-z0-9_-]+",
                    "_",
                    self.current_query,
                )
                + ".csv"
            ),
        )

        if not save:
            return
        self.exporting = True
        self.lbl_status.config(text="Exporting ALL matching rows...", fg="#fbbf24")
        query = self.current_query
        def worker():
            try:
                count = self.index.export_matches(query, save, self.current_scope)
                def finished():
                    self.exporting = False
                    self.lbl_status.config(text=f"Exported {count:,} matching rows.", fg="#34d399")
                    messagebox.showinfo("Export Complete", f"{count:,} matching rows exported.\n\n{save}")
                self.root.after(0, finished)
            except Exception as exc:
                error = str(exc)
                def failed():
                    self.exporting = False
                    messagebox.showerror("Export Error", error)
                self.root.after(0, failed)
        threading.Thread(target=worker, daemon=True).start()


def main():
    root = tk.Tk()
    app = App(
        root
    )

    def on_close():
        try:
            app.cancel_scan.set()
            app.index.close()
        except Exception:
            pass

        root.destroy()

    root.protocol(
        "WM_DELETE_WINDOW",
        on_close,
    )

    root.mainloop()


if __name__ == "__main__":
    main()

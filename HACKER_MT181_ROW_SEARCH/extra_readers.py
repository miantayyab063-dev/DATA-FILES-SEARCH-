"""Additional searchable text formats and message bodies."""

import zipfile
import shutil
import tempfile
from email import policy
from email.parser import BytesParser
from html.parser import HTMLParser
from pathlib import Path
from xml.etree import ElementTree as ET


class VisibleText(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.hidden = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "noscript"):
            self.hidden += 1
        elif tag in ("p", "br", "div", "tr", "li"):
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript"):
            self.hidden = max(0, self.hidden - 1)
        elif tag in ("p", "div", "tr", "li"):
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.hidden:
            self.parts.append(data)

    def text(self):
        return "".join(self.parts)


def visible_html(raw):
    parser = VisibleText()
    parser.feed(raw)
    return parser.text()


def read_eml(path):
    with path.open("rb") as stream:
        msg = BytesParser(policy=policy.default).parse(stream)
    headers = [(key, str(msg.get(key, ""))) for key in ("Subject", "From", "To", "Date", "Cc")]
    if any(value for _, value in headers):
        yield "Email headers", 1, [(key, value) for key, value in headers if value]
    part_no = 0
    for part in msg.walk():
        if part.is_multipart() or part.get_content_disposition() == "attachment":
            continue
        if part.get_content_type() not in ("text/plain", "text/html"):
            continue
        part_no += 1
        try:
            content = part.get_content()
        except (LookupError, UnicodeError):
            content = (part.get_payload(decode=True) or b"").decode("utf-8", "replace")
        if part.get_content_type() == "text/html":
            content = visible_html(content)
        for line_no, line in enumerate(content.splitlines(), 1):
            if line.strip():
                yield f"Email body {part_no}", line_no, [("Text", line.strip())]


def read_pptx(path):
    from pptx import Presentation

    slides = Presentation(path)
    for slide_no, slide in enumerate(slides.slides, 1):
        shape_no = 0
        for shape in slide.shapes:
            shape_no += 1
            if shape.has_text_frame:
                for line_no, paragraph in enumerate(shape.text_frame.paragraphs, 1):
                    if paragraph.text.strip():
                        yield f"Slide {slide_no} / shape {shape_no}", line_no, [("Text", paragraph.text.strip())]
            if shape.has_table:
                for row_no, row in enumerate(shape.table.rows, 1):
                    cells = [(f"Col_{i}", cell.text) for i, cell in enumerate(row.cells, 1) if cell.text]
                    if cells:
                        yield f"Slide {slide_no} / table {shape_no}", row_no, cells
        if slide.has_notes_slide:
            note = slide.notes_slide.notes_text_frame
            if note and note.text.strip():
                yield f"Slide {slide_no} / notes", 1, [("Text", note.text.strip())]


def read_odt(path):
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("content.xml"))
    line = 0
    for element in root.iter():
        if element.tag.rsplit("}", 1)[-1] not in ("p", "h"):
            continue
        value = "".join(element.itertext()).strip()
        if value:
            line += 1
            yield "Document", line, [("Text", value)]


def column_name(index):
    letters = ""
    while index:
        index, digit = divmod(index - 1, 26)
        letters = chr(65 + digit) + letters
    return letters


def read_ods(path):
    table_ns = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    office_ns = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    with zipfile.ZipFile(path) as archive, archive.open("content.xml") as xml:
        sheet_name = "Sheet"
        row_number = 0
        for event, element in ET.iterparse(xml, events=("start", "end")):
            tag = element.tag.rsplit("}", 1)[-1]
            if event == "start" and element.tag == f"{{{table_ns}}}table":
                sheet_name = element.attrib.get(f"{{{table_ns}}}name", "Sheet")
                row_number = 0
            if event != "end" or element.tag != f"{{{table_ns}}}table-row":
                continue
            cells = []
            col = 1
            for cell in element:
                if cell.tag not in (f"{{{table_ns}}}table-cell", f"{{{table_ns}}}covered-table-cell"):
                    continue
                repeats = int(cell.attrib.get(f"{{{table_ns}}}number-columns-repeated", "1"))
                value = " ".join(t.strip() for t in cell.itertext() if t and t.strip())
                if not value:
                    value = cell.attrib.get(f"{{{office_ns}}}value", "")
                if value:
                    cells.extend((column_name(n), value) for n in range(col, col + repeats))
                col += repeats
            row_repeats = int(element.attrib.get(f"{{{table_ns}}}number-rows-repeated", "1"))
            if cells:
                for n in range(row_number + 1, row_number + row_repeats + 1):
                    yield sheet_name, n, cells
            row_number += row_repeats
            element.clear()


def read_rtf(path):
    try:
        from striprtf.striprtf import rtf_to_text
    except ImportError as exc:
        raise RuntimeError("RTF reader is missing from this installation") from exc
    raw = path.read_bytes().decode("latin-1")
    for line_no, line in enumerate(rtf_to_text(raw).splitlines(), 1):
        if line.strip():
            yield "RTF", line_no, [("Text", line.strip())]


def read_archive(path, read_rows):
    with zipfile.ZipFile(path) as archive, tempfile.TemporaryDirectory() as tmp:
        for n, info in enumerate(archive.infolist(), 1):
            if info.is_dir():
                continue
            filename = info.filename.replace("\\", "/").rsplit("/", 1)[-1]
            if not filename:
                continue
            inner = Path(tmp) / f"{n}_{filename}"
            with archive.open(info) as source, inner.open("wb") as output:
                shutil.copyfileobj(source, output)
            try:
                for sheet, row_no, cells in read_rows(inner):
                    yield f"{info.filename} / {sheet}", row_no, cells
            except Exception as exc:
                raise RuntimeError(f"Archive entry {info.filename}: {exc}") from exc


def read_generic_text(path):
    with path.open("rb") as stream:
        head = stream.read(8192)
    if head.startswith((b"PK\x03\x04", b"\xd0\xcf\x11\xe0", b"\x89PNG", b"\xff\xd8\xff", b"MZ", b"7z\xbc\xaf")):
        raise RuntimeError("Unsupported binary file format; a dedicated reader or OCR is required")
    if b"\x00" in head and not head.startswith((b"\xff\xfe", b"\xfe\xff")):
        raise RuntimeError("Binary format without a text reader; convert to a supported text document")
    if head and sum(byte < 32 and byte not in (9, 10, 13) for byte in head) > len(head) // 50:
        raise RuntimeError("Unknown binary format; a dedicated reader is required")
    if head.startswith((b"\xff\xfe", b"\xfe\xff")):
        encoding = "utf-16"
    else:
        try:
            head.decode("utf-8")
            encoding = "utf-8-sig"
        except UnicodeDecodeError:
            encoding = "cp1252"
    with path.open("r", encoding=encoding, errors="replace") as stream:
        for line_no, line in enumerate(stream, 1):
            value = line.strip()
            if value:
                yield "Text", line_no, [("Text", value)]


def read_extra_rows(path, read_rows=None):
    ext = path.suffix.casefold()
    if ext == ".eml":
        yield from read_eml(path)
    elif ext == ".pptx":
        yield from read_pptx(path)
    elif ext == ".odt":
        yield from read_odt(path)
    elif ext == ".ods":
        yield from read_ods(path)
    elif ext == ".zip":
        if read_rows is None:
            raise RuntimeError("ZIP archive reader is unavailable")
        yield from read_archive(path, read_rows)
    elif ext == ".rtf":
        yield from read_rtf(path)
    elif ext in (".html", ".htm"):
        for n, line in enumerate(visible_html(path.read_text(encoding="utf-8", errors="replace")).splitlines(), 1):
            if line.strip():
                yield "HTML", n, [("Text", line.strip())]
    elif ext in (".csv", ".xlsx", ".xls", ".xlsb", ".xlsm", ".docx", ".pdf", ".txt"):
        raise RuntimeError(f"Existing reader should handle {ext}")
    else:
        yield from read_generic_text(path)

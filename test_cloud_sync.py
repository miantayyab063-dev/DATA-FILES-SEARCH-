"""Exercise paginated Gmail/Drive mirroring through the real file index."""
import base64
import io
from email.message import EmailMessage
from pathlib import Path
import sys
import tempfile
import types

sys.path.insert(0, str(Path(__file__).parent / "HACKER_MT181_ROW_SEARCH"))
from app import FastIndex
from google_sync import GoogleWorkspace


class Request:
    def __init__(self, value):
        self.value = value

    def execute(self):
        return self.value


class FakeDownload:
    def __init__(self, output, request, chunksize):
        self.output = output
        self.request = request

    def next_chunk(self, num_retries):
        self.output.write(self.request.value)
        return None, True


sys.modules["googleapiclient"] = types.ModuleType("googleapiclient")
http = types.ModuleType("googleapiclient.http")
http.MediaIoBaseDownload = FakeDownload
sys.modules["googleapiclient.http"] = http


class Gmail:
    def __init__(self):
        msg = EmailMessage()
        msg["Subject"] = "Payroll 2026"
        msg["From"] = "hr@example.test"
        msg.set_content("Employee telephone 03001234567\n")
        msg.add_attachment("Name,Number\nAisha,99887766\n".encode(),
                           maintype="text", subtype="csv", filename="report.csv")
        self.raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        self.downloads = 0

    def users(self):
        return self

    def messages(self):
        return self

    def list(self, pageToken=None, **kwargs):
        assert kwargs["maxResults"] == 500 and kwargs["includeSpamTrash"]
        if not pageToken:
            return Request({"messages": [{"id": "m1"}], "nextPageToken": "next"})
        return Request({"messages": [{"id": "m2"}]})

    def get(self, id, format, **kwargs):
        self.downloads += 1
        return Request({"raw": self.raw})


class Drive:
    def __init__(self):
        self.version = "1"
        self.downloads = 0

    def drives(self):
        return self

    def files(self):
        return self

    def list(self, pageToken=None, **kwargs):
        if kwargs.get("fields", "").startswith("nextPageToken,drives"):
            return Request({"drives": []})
        if not pageToken:
            return Request({"files": [{"id": "d1", "name": "notes.txt", "mimeType": "text/plain",
                                       "modifiedTime": self.version}], "nextPageToken": "next"})
        return Request({"files": [{"id": "sheet1", "name": "Large Sheet",
                                   "mimeType": "application/vnd.google-apps.spreadsheet",
                                   "modifiedTime": "1"}]})

    def get_media(self, fileId, **kwargs):
        self.downloads += 1
        return Request(("Direct from Drive 11223344" if self.version == "1"
                        else "Updated from Drive 55667788").encode())

    def export_media(self, **kwargs):
        raise RuntimeError("Drive export larger than 10 MB")


class Sheets:
    def spreadsheets(self):
        return self

    def values(self):
        return self

    def get(self, **kwargs):
        if "fields" in kwargs:
            return Request({"sheets": [{"properties": {"title": "Records",
                        "gridProperties": {"rowCount": 200001, "columnCount": 2}}}]})
        return Request({"values": [["Number", "77889900"]]
                        if "A200001:" in kwargs.get("range", "") else []})


with tempfile.TemporaryDirectory() as directory:
    root = Path(directory)
    data = root / "data"
    sync = GoogleWorkspace(root / "state", data)
    sync.email = "test@example.com"
    sync.gmail = Gmail()
    sync.drive = Drive()
    sync.sheets = Sheets()
    result = sync.sync()
    assert result["counts"] == {"gmail": 2, "drive": 2} and not result["errors"], result
    index = FastIndex(root / "cache.db")
    scan = index.scan(data)
    assert scan["failed"] == 0 and scan["files"] == 6, scan
    for number in ("03001234567", "99887766", "11223344", "77889900"):
        rows, _ = index.search_page(number)
        assert rows and any(number in row["full"] for row in rows), number
    sheet_row = index.search_page("77889900")[0][0]
    assert sheet_row["row"] == 200001, sheet_row
    index.search_page("11223344")
    assert index.conn.execute("SELECT COUNT(*) FROM search_cache").fetchone()[0] > 0
    sync.sync()
    assert sync.gmail.downloads == 2 and sync.drive.downloads == 1, "Unchanged cloud items redownloaded"
    sync.drive.version = "2"
    sync.sync()
    index.scan(data)
    assert not index.search_page("11223344")[0]
    assert index.search_page("55667788")[0]
    index.close()

print("PASS: paginated Gmail + attachment, Drive + 200001-row Sheets fallback, real search, incremental sync and cache invalidation")

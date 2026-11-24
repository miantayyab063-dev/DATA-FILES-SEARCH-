"""Regression: searchable indexed rows while loading, then refresh final results."""
from unittest.mock import Mock, patch
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent / "HACKER_MT181_ROW_SEARCH"))
from app import App


class Entry:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def delete(self, *_):
        self.value = ""

    def insert(self, _position, value):
        self.value = value


class FakeThread:
    launched = []

    def __init__(self, **kwargs):
        self.kwargs = kwargs

    def start(self):
        self.launched.append(self.kwargs)


class Progress(dict):
    def stop(self):
        pass


ui = Mock()
ui.entry = Entry("03001234567")
ui.scanning = True
ui.search_token = 0
ui.current_query = ""
ui.pending_query = ""
ui.search_after = None
ui.cloud_waiting = False
ui.cloud_syncing = False
ui.sync_needs_scan = False
ui.last_cloud_failed = 0
ui.index.previous_query.return_value = None
ui.page_starts = [0]
ui.page_index = 0
ui.last_scan_failed = 0
ui.start_search = lambda mode="exact": App.start_search(ui, mode)
ui.scan_started_at = 0
ui.progress = Progress()
FakeThread.launched = []

with patch("app.threading.Thread", FakeThread):
    App.start_search(ui)
    assert ui.pending_query == "03001234567"
    assert len(FakeThread.launched) == 1, "Indexed rows must be searchable while scanning"
    assert FakeThread.launched[0]["args"][0] == "03001234567"
    App.scan_done(ui, {"ok_files": 2, "files": 2, "rows": 400001,
                       "failed": 0, "errors": []}, 12.0)
    assert len(FakeThread.launched) == 2
    assert FakeThread.launched[-1]["args"][0] == "03001234567"
    assert not ui.pending_query

print("PASS: partial indexed rows searched during scan and refreshed after full scan")

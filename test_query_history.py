"""Unique-number history persists first search, last search and repeat count."""
import tempfile
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent / "HACKER_MT181_ROW_SEARCH"))
from app import FastIndex

with tempfile.TemporaryDirectory() as temp:
    db = Path(temp) / "index.db"
    index = FastIndex(db)
    first = index.record_query("001234", "7")
    repeated = index.record_query("001234", "9")
    assert first[2] == 1 and repeated[2] == 2
    assert repeated[0] == first[0] and repeated[3] == "9"
    index.close()
    reopened = FastIndex(db)
    assert reopened.previous_query("001234")[2] == 2
    rows, total = reopened.list_checked_numbers(0, 300)
    assert total == 1 and rows[0][0] == "001234"
    reopened.close()

print("PASS: every checked number retains first/last search and repeat count across restarts")

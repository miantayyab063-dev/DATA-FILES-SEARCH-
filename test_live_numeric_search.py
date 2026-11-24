"""Checks live ranking and complete rows without a display server."""
import csv
import pathlib
import sys
import tempfile
import time
from types import SimpleNamespace
from unittest.mock import Mock

sys.path.insert(0, str(pathlib.Path("HACKER_MT181_ROW_SEARCH").resolve()))
from app import App, FastIndex, numeric_match_rank


with tempfile.TemporaryDirectory() as tmp:
    directory = pathlib.Path(tmp)
    data = directory / "data"
    data.mkdir()
    with (data / "one.csv").open("w", newline="") as output:
        writer = csv.writer(output)
        writer.writerow(["person", "phone", "city"])
        writer.writerows([
            ["A", "A123456789B", "Multan"],
            ["B", "325991", "Lahore"],
            ["C", "77012345678X", "Peshawar"],
            ["D", "1234567", "Karachi"],
            ["E", "3456", "Quetta"],
        ])
    with (data / "two.csv").open("w", newline="") as output:
        writer = csv.writer(output)
        writer.writerow(["person", "phone", "city"])
        writer.writerow(["F", "999325000", "Islamabad"])

    index = FastIndex(directory / "index.db")
    assert index.scan(data)["failed"] == 0
    for term in ("1", "325"):
        matches, _ = index.search_page(term)
        assert matches and all(any(term in value for _, value in row["cells"])
                               for row in matches), term
    assert {pathlib.Path(row["path"]).name for row in index.search_page("325")[0]} == {
        "one.csv", "two.csv"
    }
    assert numeric_match_rank([("X", "A123456789B")], "123456789") == 9
    assert numeric_match_rank([("X", "012345678X")], "123456789") == 8
    assert numeric_match_rank([("X", "A1234567B")], "123456789") == 7
    assert numeric_match_rank([("X", "3456")], "123456789") == 4
    assert numeric_match_rank([("X", "numbers")], "123456789") == 0
    cursor = None
    ranks, paths, ids = [], [], []
    while True:
        page, more, cursor = index.search_ranked_page("123456789", cursor, page_size=2)
        ranks.extend(row["rank"] for row in page)
        paths.extend(row["path"] for row in page)
        ids.extend(row["id"] for row in page)
        assert all(len(row["cells"]) == 3 and row["full"] for row in page)
        if not more:
            break
    assert ranks == sorted(ranks, reverse=True), ranks
    assert ranks[0] == 9 and 8 in ranks and 7 in ranks and 4 in ranks, ranks
    assert len(ids) == len(set(ids))
    assert [row["id"] for row in index.search_ranked_page("1")[0]] == [
        row["id"] for row in index.search_page("1")[0]
    ]
    assert index.search_ranked_page("1234567890123456789")[0] == []

    ui = Mock()
    ui.search_token = 0
    ui.search_after = None
    ui.root.after.side_effect = lambda _delay, callback, *args: callback(*args)
    # Debounced entry event invokes preview; the Search button uses exact mode.
    App.on_typing(ui, SimpleNamespace(keysym="9"))
    ui.start_search.assert_called_once_with("preview")
    ui.search_done.reset_mock()
    ui.scanning = ui.cloud_syncing = False
    ui.index = index
    App.search_worker(ui, "123456789", ui.search_token, time.perf_counter(),
                      0, (9, 0), "preview")
    assert ui.search_done.call_count == 2  # exact first, ranked results next
    assert [row["rank"] for row in ui.search_done.call_args.args[2]][:1] == [9]
    ui.search_done.reset_mock()
    App.search_worker(ui, "325", ui.search_token, time.perf_counter(),
                      0, 0, "exact")
    assert ui.search_done.call_count == 1
    assert all(row["rank"] == 3 for row in ui.search_done.call_args.args[2])
    index.close()
print("PASS: digits 1/325 in complete rows; 9,8,7,...4 ranking; two files; paging")

"""Display-backed Tk smoke test for both search modes and visible side panels."""
import pathlib
import tempfile
import time
import tkinter as tk

import sys
sys.path.insert(0, str(pathlib.Path("HACKER_MT181_ROW_SEARCH").resolve()))
import app

with tempfile.TemporaryDirectory() as directory:
    work = pathlib.Path(directory)
    (work / "data").mkdir()
    (work / "data" / "TeamFile.csv").write_text(
        "Name,Phone,City\nAisha,03250000,Lahore\n", encoding="utf-8")
    app.__file__ = str(work / "app.py")
    root = tk.Tk()
    gui = app.App(root)
    deadline = time.monotonic() + 8
    while (gui.scanning or not gui.index.stats()[0]) and time.monotonic() < deadline:
        root.update()
        time.sleep(0.02)
    root.update()
    assert gui.index.stats()[0] == 1
    gui.search_scope.set("file")
    gui.entry.insert(0, "TeamFile")
    gui.start_search("exact")
    while not gui.current_results and time.monotonic() < deadline:
        root.update()
        time.sleep(0.02)
    assert len(gui.current_results) == 1 and gui.current_results[0]["sheet"] == "File"
    gui.entry.delete(0, "end")
    gui.entry.insert(0, "Aisha")
    gui.search_scope.set("data")
    gui.start_search("exact")
    while (not gui.current_results or gui.current_results[0]["sheet"] == "File") and time.monotonic() < deadline:
        root.update()
        time.sleep(0.02)
    assert len(gui.current_results) == 1 and "03250000" in gui.current_results[0]["full"]
    assert gui.tree.winfo_width() > 400 and gui.history_tree.winfo_height() > 100
    assert gui.log_text.winfo_height() > 100 and gui.history_tree.winfo_ismapped()
    assert gui.log_text.winfo_ismapped()
    gui.cancel_scan.set()
    gui.index.close()
    root.destroy()

print("PASS: Tk UI file/content modes, results, visible history and activity logs")

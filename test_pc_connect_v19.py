"""Integration checks for local-drive discovery and separate search modes."""
import csv
import pathlib
import sys
import tempfile
from docx import Document
from openpyxl import Workbook
from reportlab.pdfgen import canvas

sys.path.insert(0, str(pathlib.Path("HACKER_MT181_ROW_SEARCH").resolve()))
from app import FastIndex, pc_files_under


with tempfile.TemporaryDirectory() as temp:
    work = pathlib.Path(temp)
    c = work / "C"
    d = work / "D"
    documents = c / "Users" / "Mian" / "Documents"
    downloads = d / "Downloads"
    app_home = c / "InstalledApp"
    for path in (documents, downloads, app_home / "data", c / "Windows",
                 c / "Users" / "Mian" / "AppData", d / "cache"):
        path.mkdir(parents=True)
    with (documents / "CustomerReport.csv").open("w", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(["Person", "Number", "City"])
        writer.writerow(["Mian", "AA-325-XX", "Karachi"])
    (downloads / "MeetingNotes.txt").write_text("Client 123456789 confirmed.\n", encoding="utf-8")
    doc = Document()
    doc.add_paragraph("Project Garrison key 554433")
    doc.save(documents / "Offer.docx")
    sheet = Workbook()
    sheet.active.append(["Employee", "Phone", "Office"])
    sheet.active.append(["Talha", "7722001", "Lahore"])
    sheet.save(downloads / "Team.xlsx")
    pdf = canvas.Canvas(str(downloads / "Agreement.pdf"))
    pdf.drawString(50, 750, "Signed amount 889977")
    pdf.save()
    (app_home / "data" / "Source.csv").write_text("name,phone\nAli,99001\n", encoding="utf-8")
    for skipped in (c / "Windows" / "Secret.txt", d / "cache" / "Hidden.csv",
                    c / "Users" / "Mian" / "AppData" / "private.docx"):
        skipped.write_text("SHOULD NOT BE INDEXED 738921", encoding="utf-8")
    (downloads / "program.exe").write_bytes(b"123456789")
    (app_home / "app.txt").write_text("SHOULD NOT BE INDEXED 738921", encoding="utf-8")

    roots = [c, d, app_home / "data"]
    found = pc_files_under(roots, exclude_dirs=[app_home])
    assert {p.name for p in found} == {"CustomerReport.csv", "MeetingNotes.txt", "Source.csv", "Offer.docx", "Team.xlsx", "Agreement.pdf"}, found
    db = FastIndex(work / "index.db")
    try:
        result = db.scan(roots, pc_mode=True, exclude_dirs=[app_home])
        assert result["ok_files"] == 6 and result["failed"] == 0, result
        for number, suffix in (("554433", "Offer.docx"), ("7722001", "Team.xlsx"),
                               ("889977", "Agreement.pdf")):
            assert db.search_page(number)[0][0]["path"].endswith(suffix)
        content, more = db.search_page("325")
        assert not more and len(content) == 1 and content[0]["cells"][-1][1] == "Karachi"
        assert {p["path"] for p in db.search_page("Mian")[0]} == {str(documents / "CustomerReport.csv")}
        assert len(db.search_page("738921")[0]) == 0
        file_hits, more = db.search_files_page("customerreport")
        assert not more and len(file_hits) == 1
        assert file_hits[0]["row"] == "—" and file_hits[0]["cells"][1][1].endswith("CustomerReport.csv")
        assert db.search_files_page("325")[0] == []
        assert {pathlib.Path(p["path"]).name for p in db.search_files_page("downloads")[0]} == {
            "MeetingNotes.txt", "Team.xlsx", "Agreement.pdf"}
        assert db.search_ranked_page("Karachi")[0][0]["full"].endswith("[C]: Karachi")
        assert len(db.search_ranked_page("123456789")[0]) == 1
        assert db.scan(roots, pc_mode=True, exclude_dirs=[app_home])["ok_files"] == 6
        (downloads / "MeetingNotes.txt").unlink()
        assert db.scan(roots, pc_mode=True, exclude_dirs=[app_home])["ok_files"] == 5
        assert db.search_files_page("MeetingNotes")[0] == []
    finally:
        db.close()

print("PASS: C/D document discovery, system/cache exclusion, content and filename modes, rescan")

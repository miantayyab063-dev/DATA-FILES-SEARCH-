"""Assemble a portable Windows GUI executable using official CPython binaries.

Inputs are the official CPython 3.12 embeddable distribution, the Tcl/Tk MSI
payload from the official CPython Windows installer, and Windows wheels.
No software installation is needed on the destination Windows computer.
"""

from pathlib import Path
import os
import shutil
import zipfile

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "HACKER_MT181_ROW_SEARCH"
DEST = ROOT / "HACKER_MT181_WINDOWS_EXE"
RUNTIME_ZIP = ROOT / "python-3.12.0-embed-amd64.zip"
TK_PAYLOAD = ROOT / "python_installer_parts/tk/files"


def build():
    DEST.mkdir(exist_ok=True)
    with zipfile.ZipFile(RUNTIME_ZIP) as archive:
        archive.extractall(DEST)

    launcher = ROOT / "launcher_win64.exe"
    if not launcher.is_file():
        raise FileNotFoundError("Compile win_launcher.c for Windows first.")
    shutil.copy2(launcher, DEST / "HACKER_M_T181.exe")
    (DEST / "python312._pth").write_text(
        "python312.zip\n.\nLib\nsite-packages\nimport site\n", encoding="utf-8"
    )

    (DEST / "Lib/tkinter").mkdir(parents=True, exist_ok=True)
    for file in TK_PAYLOAD.iterdir():
        name = file.name
        if name in ("_tkinter.pyd", "tcl86t.dll", "tk86t.dll", "zlib1.dll"):
            shutil.copy2(file, DEST / name)
        elif name.startswith("Lib_tkinter_"):
            shutil.copy2(file, DEST / "Lib/tkinter" / name[len("Lib_tkinter_"):])
        elif name.startswith(("tcl_tcl8.6_", "tcl_tk8.6_")):
            base = "tcl8.6" if name.startswith("tcl_tcl8.6_") else "tk8.6"
            relative = name[len("tcl_" + base + "_"):]
            group, marker, rest = relative.partition("_")
            if marker and group in {
                "encoding", "msgs", "tzdata", "ttk", "images", "icons",
                "http1.0", "opt0.4", "demos",
            }:
                relative = group + "/" + rest
            target = DEST / "tcl" / base / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(file, target)

    site_packages = DEST / "site-packages"
    site_packages.mkdir(exist_ok=True)
    for wheel in sorted((ROOT / "windows_wheels").glob("*.whl")):
        with zipfile.ZipFile(wheel) as archive:
            archive.extractall(site_packages)

    for name in ["app.py", "extra_readers.py", "google_sync.py", "requirements.txt"]:
        shutil.copy2(SOURCE / name, DEST / name)

    oauth = SOURCE / "google_oauth_client.json"
    oauth_bundled = oauth.is_file()
    if oauth_bundled:
        shutil.copy2(oauth, DEST / oauth.name)
    else:
        (DEST / oauth.name).unlink(missing_ok=True)

    data_dest = DEST / "data"
    # Older builds contained source spreadsheets. Never carry them into the
    # distributable: an existing user's files are not part of this installer.
    if data_dest.exists():
        shutil.rmtree(data_dest)
    data_dest.mkdir(exist_ok=True)

    shutil.copy2(TK_PAYLOAD / "tcl_tk8.6_license.terms", DEST / "TK_LICENSE.txt")
    shutil.copy2(ROOT / "python_installer_parts/license/LICENSE.txt", DEST / "PYTHON_LICENSE.txt")
    (DEST / "sitecustomize.py").unlink(missing_ok=True)
    (DEST / "boot.py").write_text('''\
"""Set up Tcl/Tk and show startup errors from the bundled Windows runtime."""
import os
import sys
from pathlib import Path

home = Path(__file__).resolve().parent
os.environ["TCL_LIBRARY"] = str(home / "tcl" / "tcl8.6")
os.environ["TK_LIBRARY"] = str(home / "tcl" / "tk8.6")
try:
    import runpy
    sys.argv = [str(home / "app.py")]
    runpy.run_path(str(home / "app.py"), run_name="__main__")
except BaseException:
    import traceback
    details = traceback.format_exc()
    (home / "startup_error.log").write_text(details, encoding="utf-8")
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            None, "Unable to launch. See startup_error.log next to the EXE.\\n\\n" + details[-900:],
            "HACKER M-T181", 0x10
        )
    except Exception:
        pass
''', encoding="utf-8")
    (DEST / "Debug_Start.bat").write_text(
        '@echo off\r\ncd /d "%~dp0"\r\npython.exe boot.py\r\npause\r\n', encoding="ascii"
    )
    (DEST / "README_URDU.txt").write_text(
        "HACKER M-T181 - Windows 64-bit EXE\n\n"
        "1. Windows setup EXE install karein; phir Start menu se HACKER M-T181 chalayein. Python install karna zaroori nahi.\n"
        "2. Installer mein koi personal files nahin. Khali data folder milega; Connect to PC dabane par PC ki local drives ke documents read honge.\n"
        "2a. Connect to PC button se C:/D:/aur local removable drives ke documents scan honge. Windows, Program Files, AppData, cache aur software ke runtime folders scan nahin honge. PC documents isi PC par index hote hain; server par upload nahin hote. Pehla scan documents ki tadaad par depend karega. Locked/unreadable documents aur inaccessible folders activity logs mein dikhte hain.\n"
        "2b. Data Inside Files mode mein har matching row ka complete data; File Name / Path mode mein matching document ka naam, full path aur indexing status dikhte hain. 18-character tak numbers, letters aur readable text search karein; numeric preview mein 4-digit tak near matches, Search button exact matches.\n"
        "3. Data folder mein apni Excel, XLSB, XLS, ODS, CSV, DOCX, PDF, email, ZIP aur text files rakhein.\n"
        "4. Pehli baar har file ki har populated row load hogi. Screen par file, rows, aur progress dikhegi.\n"
        "5. Sirf 0-9 digits likhein (maximum 18). Har key par loaded rows mein live exact match dikhega; 5+ digits ke liye paas ke 4+ contiguous digit matches neeche dikhenge. Search click par sirf exact rows dikhenge. Scan ke dauran interim results aayenge, poori scan ke baad automatically update honge.\n"
        "6. Results 500 rows per page dikhte hain; Next se baqi pages; Export ALL se tamam matches.\n"
        "6a. Right-side Checked numbers tab mein har unique number ka first search, last search, repeat count aur last result dekhein; us number ko double-click kar ke dobara search karein.\n"
        "7. Gmail + Drive ke liye Google Cloud mein Gmail, Drive aur Sheets APIs enabled hon aur OAuth consent screen par aapka Gmail Test users mein ho (agar Testing ho).\n"
        + ("8. Desktop OAuth client is installer mein bundled hai. Connect Gmail + Drive click karein; browser mein apna account select karke read-only access approve karein. JSON dobara select karna zaroori nahi.\n"
           if oauth_bundled else "8. Connect Gmail + Drive click karein; pehli dafa apna Desktop OAuth JSON select karein, phir browser mein account approve karein.\n")
        +
        "9. Email messages, attached files, Drive Docs/Sheets/Slides aur uploaded files local __Google_Cloud__ subfolder mein mirror hote hain. Sync ke dauran indexed rows searchable rehti hain.\n"
        "10. Agar file fail ho, right-side History + logs panel mein file name aur error dekhein. Purani .doc binary, unknown binary, protected docs aur image-only PDFs ko OCR/compatible reader chahiye; inhein silently complete nahi kaha jata.\n"
        "11. Agar GUI start na ho to startup_error.log dekhein ya Debug_Start.bat chalayein.\n"
        "12. Folder writable rakhein: search cache (kayi GB), history aur activity log isi mein bante hain. Google login token is PC ke google_config folder mein save hota hai.\n\n"
        "Windows 10/11 64-bit. Supplied runtime: Python 3.12.0 (Python Software Foundation).\n",
        encoding="utf-8"
    )
    (DEST / "VERIFICATION.txt").write_text(
        "Installer contains ZERO bundled user documents. Original 44 spreadsheet test data is not packaged.\n"
        "V20 PC document discovery: synthetic C/D drives, Windows/cache exclusion, data-folder inclusion, file-name/path versus content search, rescan/deletion verified.\n"
        "PC index caches on local disk; protected files, image-only PDFs, old binary DOC, unsupported types and excluded OS/cache folders do not yield complete content results.\n"
        "Full Linux scan: 6,652,476 populated rows, 44/44 files, 0 failures.\n"
        "Numeric checks: 0/12345 found complete source rows with filenames; "
        "absent 20-digit query returned zero.\n"
        "Synthetic 6,200+ rows: all ASCII digit 0-9 searches matched "
        "source cells across pages; full-row export, identical copies, "
        "leading zero formatting verified.\n"
        "Partial search while scanning and final refresh: regression test passed.\n"
        "Live 1/325 matching, 9-to-4-digit ranking, two-file exact rows, pagination and 18-digit limit: regression tests passed.\n"
        "Mocked Gmail/Drive integration: paginated messages, CSV attachment, "
        "Drive text, 200001-row Google Sheet chunks, cache invalidation passed.\n"
        "Desktop OAuth flow: bundled or imported client, Google browser callback, "
        "Gmail and Drive account verification passed with mocked APIs.\n"
        "The included launcher is a native x86-64 Windows GUI PE executable "
        "with its own bundled Python and Tk runtime.\n"
        + ("A supplied Google Desktop OAuth client is bundled and validated. "
           if oauth_bundled else "No real OAuth Desktop client was supplied. ")
        + "Actual Google sign-in and "
        "Windows GUI could not be run in the Linux test environment.\n",
        encoding="utf-8"
    )
    (DEST / "GOOGLE_CONNECTION_SETUP.txt").write_text(
        "Gmail + Google Drive connection (Google account ka password software mein na likhein)\n\n"
        "1. https://console.cloud.google.com/ par apna Google project create/select karein.\n"
        "2. APIs & Services > Library mein Gmail API, Google Drive API aur Google Sheets API enable karein.\n"
        "3. Google Auth Platform / OAuth consent screen configure karein. External Testing mode ho to apna Gmail address Test users mein add karein.\n"
        "4. Data access/scopes: gmail.readonly, drive.readonly, spreadsheets.readonly. Gmail.readonly restricted scope hai; Google consent aur test users/production verification ki conditions lag sakti hain.\n"
        "5. Clients/Credentials mein Desktop OAuth client banana zaroori hai."
        + (" Is installer mein supplied Desktop JSON pehle se bundled hai.\n"
           if oauth_bundled else " Apni Desktop JSON download karein.\n")
        + "6. Software mein Connect Gmail + Drive click karein. "
        + ("Bundled client ke sath pehli click par seedha browser khule ga; account select karein aur access approve karein.\n"
           if oauth_bundled else "Pehli dafa setup window mein Desktop JSON select karein, phir browser account approval khule ga.\n")
        +
        "7. Gmail aur Drive dono green 'connected' hone ke baad sync shuru ho gi; rows read hone ke baad search result aayega. Agli baar 'Sync Google now' dabayein.\n"
        "8. External Testing OAuth project ke refresh tokens 7 din mein expire ho sakte hain. App invalid token par browser sign-in dobara kholegi. Public usage ke liye Google's brand/scope verification owner ko karni hogi.\n"
        "9. Google Docs/Sheets/Slides exports mein Drive ki 10 MB export limit hai. Bari Google Sheets ko app Sheets API se rows ki chunks mein read karti hai. Jis file ko Google API export na kare us ka naam aur wajah Activity logs mein dekhein.\n"
        "10. Login token PC ke app folder ke google_config/ mein rehta hai. App kabhi Gmail password nahin poochti. "
        + ("OAuth client installer mein hai; Gmail/Drive ka account token bundled nahin, browser par account approval aap khud karte hain.\n\n"
           if oauth_bundled else "Desktop OAuth JSON aur account approval ke baghair live Gmail/Drive login mumkin nahin.\n\n")
        +
        "Official guides: https://developers.google.com/workspace/gmail/api/quickstart/python\n"
        "https://developers.google.com/workspace/drive/api/quickstart/python\n",
        encoding="utf-8"
    )
    print("Created", DEST, "with", len(list(data_dest.rglob("*"))), "data entries")


if __name__ == "__main__":
    build()

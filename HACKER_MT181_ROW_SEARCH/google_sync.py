"""Read-only Gmail and Drive mirroring for the local row search index.

Every fetched item is saved beneath the selected data folder. This allows the
same file parser, SQLite cache, pagination, and complete-row search to handle
local files, mail attachments, and Drive documents identically.
"""

import base64
import csv
import json
import os
import re
from email import policy
from email.parser import BytesParser
from pathlib import Path


SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]

EXPORTS = {
    "application/vnd.google-apps.document": ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"),
    "application/vnd.google-apps.spreadsheet": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    "application/vnd.google-apps.presentation": ("application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx"),
    "application/vnd.google-apps.drawing": ("application/pdf", ".pdf"),
}


def safe_filename(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name or "Untitled")
    return name.rstrip(" .")[:100] or "Untitled"


def write_atomic(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".partial")
    with tmp.open("wb") as out:
        out.write(data)
    os.replace(tmp, path)


def save_json(path, data):
    write_atomic(path, json.dumps(data, ensure_ascii=False).encode("utf-8"))


def remove_saved(folder, relative_paths):
    for relative in relative_paths:
        candidate = folder / relative
        if candidate.is_file():
            candidate.unlink()


class GoogleWorkspace:
    def __init__(self, state_dir, data_folder, notify=lambda _text: None):
        self.state_dir = Path(state_dir)
        self.data_folder = Path(data_folder)
        self.state_dir.mkdir(parents=True, exist_ok=True)
        self.client_path = self.state_dir / "google_oauth_client.json"
        # A provider-owned Desktop OAuth client can be supplied beside app.py
        # when packaging a private installer. Never synthesize an OAuth ID.
        self.bundled_client_path = Path(__file__).resolve().parent / "google_oauth_client.json"
        self.token_path = self.state_dir / "google_oauth_token.json"
        self.state_path = self.state_dir / "google_sync_state.json"
        self.notify = notify
        self.credentials = None
        self.email = ""
        self.gmail = None
        self.drive = None
        self.sheets = None

    def client_available(self):
        return self.client_path.is_file() or self.bundled_client_path.is_file()

    def import_client(self, source):
        info = json.loads(Path(source).read_text(encoding="utf-8-sig"))
        installed = info.get("installed")
        if not isinstance(installed, dict) or not installed.get("client_id"):
            raise ValueError("Select a Google OAuth Desktop application JSON file.")
        if installed.get("auth_uri") not in {
            "https://accounts.google.com/o/oauth2/auth",
            "https://accounts.google.com/o/oauth2/v2/auth",
        } or installed.get("token_uri") not in {
            "https://oauth2.googleapis.com/token",
            "https://www.googleapis.com/oauth2/v4/token",
        }:
            raise ValueError("Invalid Google OAuth endpoints in the Desktop JSON.")
        if not installed["client_id"].endswith(".apps.googleusercontent.com"):
            raise ValueError("Invalid Desktop OAuth client ID.")
        if Path(source).resolve() != self.client_path.resolve():
            write_atomic(self.client_path, json.dumps(info).encode("utf-8"))
            self.token_path.unlink(missing_ok=True)

    def connect(self, client_json=None):
        try:
            from google.auth.transport.requests import Request
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise RuntimeError("Google API libraries are missing from this installation.") from exc

        if client_json:
            self.import_client(client_json)

        if not self.client_path.exists() and self.bundled_client_path.is_file():
            self.import_client(self.bundled_client_path)

        if not self.client_path.exists():
            raise FileNotFoundError("Select your Google OAuth Desktop client JSON to connect.")

        creds = None
        if self.token_path.exists():
            try:
                creds = Credentials.from_authorized_user_file(str(self.token_path), SCOPES)
            except (ValueError, OSError):
                self.token_path.unlink(missing_ok=True)
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as exc:
                # A revoked / testing-mode refresh token needs browser consent
                # again. Keep network/API problems visible to the user.
                from google.auth.exceptions import RefreshError
                if not isinstance(exc, RefreshError):
                    raise
                self.notify("Google session expired; opening sign-in again.")
                self.token_path.unlink(missing_ok=True)
                creds = None
        if not creds or not creds.valid or not set(SCOPES).issubset(set(creds.scopes or [])):
            self.notify("Opening Google sign-in. Approve read-only Gmail, Drive and Sheets access.")
            flow = InstalledAppFlow.from_client_secrets_file(str(self.client_path), SCOPES)
            creds = flow.run_local_server(port=0, open_browser=True,
                                          timeout_seconds=300, prompt="select_account")

        write_atomic(self.token_path, creds.to_json().encode("utf-8"))
        self.credentials = creds
        self.gmail = build("gmail", "v1", credentials=creds, cache_discovery=False)
        self.drive = build("drive", "v3", credentials=creds, cache_discovery=False)
        self.sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
        self.email = self.gmail.users().getProfile(userId="me").execute()["emailAddress"]
        drive_user = self.drive.about().get(fields="user(emailAddress)").execute()
        self.notify(f"Gmail connected: {self.email}")
        self.notify(f"Drive connected: {drive_user.get('user', {}).get('emailAddress', self.email)}")
        return self.email

    def _state(self):
        if self.state_path.exists():
            try:
                return json.loads(self.state_path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                self.notify("Cloud sync state unreadable; checking all items again.")
        return {"account": self.email, "gmail": {}, "drive": {}}

    def sync(self):
        if not self.gmail or not self.drive or not self.email:
            raise RuntimeError("Connect Google first.")
        state = self._state()
        if state.get("account") != self.email:
            state = {"account": self.email, "gmail": {}, "drive": {}}
        root = self.data_folder / "__Google_Cloud__" / safe_filename(self.email)
        errors = []
        totals = {}
        for name, operation in (("gmail", self.sync_gmail), ("drive", self.sync_drive)):
            try:
                totals[name] = operation(root, state, errors)
            except Exception as exc:
                errors.append(f"{name.upper()} sync stopped: {exc}")
                self.notify(errors[-1])
            save_json(self.state_path, state)
        return {"counts": totals, "errors": errors}

    def sync_gmail(self, root, state, errors):
        folder = root / "Gmail"
        old = state.setdefault("gmail", {})
        seen = set()
        count = 0
        page_token = None
        while True:
            request = self.gmail.users().messages().list(
                userId="me", maxResults=500, includeSpamTrash=True,
                pageToken=page_token,
            )
            page = request.execute()
            for item in page.get("messages", []):
                mid = item["id"]
                seen.add(mid)
                count += 1
                cached = old.get(mid)
                if cached and all((folder / p).is_file() for p in cached.get("paths", [])):
                    continue
                try:
                    raw_url = self.gmail.users().messages().get(
                        userId="me", id=mid, format="raw"
                    ).execute()["raw"]
                    raw = base64.urlsafe_b64decode(raw_url + "=" * (-len(raw_url) % 4))
                    mail = BytesParser(policy=policy.default).parsebytes(raw)
                    subject = safe_filename(str(mail.get("Subject", "No subject")))
                    relative = Path(mid) / (subject + ".eml")
                    write_atomic(folder / relative, raw)
                    paths = [str(relative)]
                    for n, part in enumerate(mail.walk(), 1):
                        attachment_name = part.get_filename()
                        if not attachment_name:
                            continue
                        payload = part.get_payload(decode=True)
                        if payload is None:
                            continue
                        attachment = Path(mid) / f"{n:03d}_{safe_filename(attachment_name)}"
                        write_atomic(folder / attachment, payload)
                        paths.append(str(attachment))
                    if cached:
                        remove_saved(folder, set(cached.get("paths", [])) - set(paths))
                    old[mid] = {"paths": paths}
                except Exception as exc:
                    detail = f"Gmail message {mid}: {exc}"
                    errors.append(detail)
                    self.notify(detail)
                if count % 10 == 0:
                    self.notify(f"Gmail: checked {count:,} messages; {len(old):,} mirrored")
                if count % 250 == 0:
                    save_json(self.state_path, state)
            page_token = page.get("nextPageToken")
            if not page_token:
                break
        # An API error anywhere above leaves the old files intact. Remove only
        # after a full, paginated inventory succeeds.
        for mid in set(old) - seen:
            remove_saved(folder, old.pop(mid).get("paths", []))
        self.notify(f"Gmail complete: {count:,} messages checked")
        return count

    def sync_drive(self, root, state, errors):
        from googleapiclient.http import MediaIoBaseDownload

        folder = root / "Drive"
        old = state.setdefault("drive", {})
        seen = set()
        count = 0
        sources = [{"corpora": "user"}]
        inventory_complete = True
        try:
            next_drive = None
            while True:
                shared = self.drive.drives().list(
                    pageSize=100, pageToken=next_drive, fields="nextPageToken,drives(id,name)"
                ).execute()
                sources.extend({"corpora": "drive", "driveId": d["id"]} for d in shared.get("drives", []))
                next_drive = shared.get("nextPageToken")
                if not next_drive:
                    break
        except Exception as exc:
            inventory_complete = False
            detail = f"Shared Drive inventory failed: {exc}"
            errors.append(detail)
            self.notify(detail)

        for source in sources:
          page_token = None
          while True:
            page = self.drive.files().list(
                pageSize=1000, pageToken=page_token, q="trashed = false",
                fields="nextPageToken,incompleteSearch,files(id,name,mimeType,modifiedTime,size,md5Checksum,webViewLink)",
                includeItemsFromAllDrives=True, supportsAllDrives=True, **source,
            ).execute()
            if page.get("incompleteSearch"):
                inventory_complete = False
                detail = f"Drive {source} reports incomplete search; data may be missing"
                errors.append(detail)
                self.notify(detail)
            for item in page.get("files", []):
                mime = item.get("mimeType", "")
                if mime == "application/vnd.google-apps.folder":
                    continue
                fid = item["id"]
                if fid in seen:
                    continue
                count += 1
                seen.add(fid)
                if mime.startswith("application/vnd.google-apps.") and mime not in EXPORTS:
                    detail = f"Drive {item.get('name', fid)}: Google format {mime} cannot be exported"
                    errors.append(detail)
                    self.notify(detail)
                    continue
                export = EXPORTS.get(mime)
                suffix = export[1] if export else ""
                filename = safe_filename(item.get("name", fid))
                if suffix and not filename.casefold().endswith(suffix):
                    filename += suffix
                relative = f"{fid}_{filename}"
                signature = "|".join(str(item.get(k, "")) for k in ("modifiedTime", "size", "md5Checksum", "mimeType"))
                cached = old.get(fid)
                previous_paths = (cached.get("paths", [cached.get("path")]) if cached else [])
                if cached and cached.get("signature") == signature and all(
                    p and (folder / p).is_file() for p in previous_paths
                ):
                    continue
                try:
                    try:
                        request = (self.drive.files().export_media(fileId=fid, mimeType=export[0])
                                   if export else self.drive.files().get_media(fileId=fid, supportsAllDrives=True))
                        target = folder / relative
                        target.parent.mkdir(parents=True, exist_ok=True)
                        partial = target.with_name(target.name + ".partial")
                        with partial.open("wb") as out:
                            downloader = MediaIoBaseDownload(out, request, chunksize=1024 * 1024)
                            done = False
                            while not done:
                                _, done = downloader.next_chunk(num_retries=3)
                        os.replace(partial, target)
                        new_paths = [relative]
                    except Exception:
                        if mime != "application/vnd.google-apps.spreadsheet":
                            raise
                        self.notify(f"Large Google Sheet: reading cells in chunks for {filename}")
                        new_paths = self.sync_sheet_cells(fid, filename, folder)
                    if cached:
                        remove_saved(folder, set(previous_paths) - set(new_paths))
                    old[fid] = {"paths": new_paths, "signature": signature}
                except Exception as exc:
                    partial = (folder / relative).with_name(relative + ".partial")
                    partial.unlink(missing_ok=True)
                    detail = f"Drive {item.get('name', fid)}: {exc}"
                    errors.append(detail)
                    self.notify(detail)
                if count % 10 == 0:
                    self.notify(f"Drive: checked {count:,} files; {len(old):,} mirrored")
                if count % 250 == 0:
                    save_json(self.state_path, state)
            page_token = page.get("nextPageToken")
            if not page_token:
                break
        if inventory_complete:
            for fid in set(old) - seen:
                prior = old.pop(fid)
                remove_saved(folder, prior.get("paths", [prior.get("path")]))
        self.notify(f"Drive complete: {count:,} files checked")
        return count

    def sync_sheet_cells(self, fid, filename, folder):
        """Export large Sheets beyond Drive's 10 MB export limit, all tabs."""
        metadata = self.sheets.spreadsheets().get(
            spreadsheetId=fid,
            fields="sheets(properties(title,gridProperties(rowCount,columnCount)))",
        ).execute()
        paths = []
        for n, sheet in enumerate(metadata.get("sheets", []), 1):
            prop = sheet["properties"]
            title = prop["title"]
            size = prop.get("gridProperties", {})
            max_row = int(size.get("rowCount", 0))
            max_col = max(1, int(size.get("columnCount", 1)))
            letters = ""
            while max_col:
                max_col, rem = divmod(max_col - 1, 26)
                letters = chr(65 + rem) + letters
            relative = f"{fid}_{filename}/{n:03d}_{safe_filename(title)}.csv"
            target = folder / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            partial = target.with_name(target.name + ".partial")
            try:
                with partial.open("w", newline="", encoding="utf-8-sig") as out:
                    writer = csv.writer(out)
                    for start in range(1, max_row + 1, 1000):
                        end = min(max_row, start + 999)
                        sheet_name = title.replace("'", "''")
                        values = self.sheets.spreadsheets().values().get(
                            spreadsheetId=fid,
                            range=f"'{sheet_name}'!A{start}:{letters}{end}",
                            valueRenderOption="FORMATTED_VALUE",
                        ).execute().get("values", [])
                        for offset in range(end - start + 1):
                            writer.writerow(values[offset] if offset < len(values) else [])
                        if start % 20000 == 1:
                            self.notify(f"Google Sheet {filename} / {title}: {end:,}/{max_row:,} rows checked")
                os.replace(partial, target)
            except Exception:
                partial.unlink(missing_ok=True)
                raise
            paths.append(relative)
        return paths

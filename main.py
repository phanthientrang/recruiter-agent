import asyncio
import contextlib
import json
import os
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Annotated, Optional, TypedDict

from dotenv import load_dotenv
from langchain_core.messages import SystemMessage
from langchain_core.tools import tool
from langchain_openai import ChatOpenAI
from langgraph.graph import StateGraph, START, END
from langgraph.graph.message import add_messages
from langgraph.prebuilt import ToolNode, tools_condition
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from pydantic import BaseModel

from greennode_agentbase import GreenNodeAgentBaseApp, RequestContext, PingStatus

load_dotenv()

# ---------------------------------------------------------------------------
# Constants — Skill 1 & 2
# ---------------------------------------------------------------------------
JOBS_BASE_DIR = os.environ.get("JOBS_BASE_DIR", os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.environ.get("EXCEL_PATH", "")
TA_INCHARGE = os.environ.get("TA_INCHARGE") or os.environ.get("USERNAME", "Unknown")
CV_POLL_INTERVAL = int(os.environ.get("CV_POLL_INTERVAL", "60"))
PROCESSED_FILES_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "processed_cvs.json")

JOB_SUB_FOLDERS = ["LinkedIn", "VNG Careers", "Referral", "TA Search", "Others"]

DB_COLUMNS = [
    "No", "Request code", "Candidate name", "Processed Team", "Processed Position",
    "Entry date", "Source", "Referrer", "Latest company", "Latest position",
    "Email", "Phone", "Stage", "Status", "Note", "Reason for failure/withdrawal",
    "Last drawn salary", "Expected salary (Monthly Gross)", "TA Incharge", "Profile",
]

# ---------------------------------------------------------------------------
# Constants — Skill 3
# ---------------------------------------------------------------------------
TEAM_EXCEL_PATH = os.environ.get("TEAM_EXCEL_PATH", "")
SYNC_STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sync_state.json")
CONFLICT_SHEET = "Conflict Log"

# Schedule: TA_INCHARGE value (case-insensitive) → (hour, minute)
SYNC_SCHEDULE: dict[str, tuple[int, int]] = {
    "trangptt12": (7, 30),
    "hautt2":     (8,  0),
    "huyenplt":   (8, 30),
    "nhihm":      (9,  0),
}

# Columns synced to team DB — Profile excluded (personal file path, not a shared asset)
SYNC_COLUMNS = [c for c in DB_COLUMNS if c not in ("No", "Profile")]

CONFLICT_LOG_COLUMNS = [
    "Timestamp", "Email", "Request code", "Candidate name",
    "Field", "My value", "Team value", "My TA", "Team TA",
]

# In-memory alert store
_alerts: list[dict] = []


# ---------------------------------------------------------------------------
# Shared Excel helpers
# ---------------------------------------------------------------------------
def _open_wb(path: str, read_only: bool = False):
    keep_vba = Path(path).suffix.lower() == ".xlsm"
    return load_workbook(path, read_only=read_only, keep_vba=keep_vba)


def _get_headers(ws) -> list[str]:
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        return [str(c).strip() if c is not None else "" for c in row]
    return []


# ---------------------------------------------------------------------------
# Skill 2 — CV text extraction
# ---------------------------------------------------------------------------
def _resolve_path(file_path: str) -> str:
    """Resolve the actual file path, handling Unicode normalization differences
    (NFC vs NFD) that occur with Vietnamese and other accented filenames on Windows."""
    file_path = os.path.normpath(file_path)
    if os.path.exists(file_path):
        return file_path
    # Try NFC (Windows filesystem norm)
    nfc = unicodedata.normalize("NFC", file_path)
    if os.path.exists(nfc):
        return nfc
    # Scan parent directory for a name that matches after NFC normalization
    parent, name = os.path.dirname(file_path), os.path.basename(file_path)
    name_nfc = unicodedata.normalize("NFC", name).lower()
    if os.path.isdir(parent):
        for fname in os.listdir(parent):
            if unicodedata.normalize("NFC", fname).lower() == name_nfc:
                return os.path.join(parent, fname)
    return file_path  # return as-is; caller will get a clear error


def _extract_pdf_text(file_path: str) -> str:
    import pypdf
    # Open via Python's built-in open() so Windows Unicode API handles the path,
    # then pass the file object — avoids encoding issues inside pypdf itself.
    with open(file_path, "rb") as f:
        reader = pypdf.PdfReader(f)
        return "\n".join(page.extract_text() or "" for page in reader.pages)


def _extract_docx_text(file_path: str) -> str:
    from docx import Document
    with open(file_path, "rb") as f:
        doc = Document(f)
    return "\n".join(p.text for p in doc.paragraphs)


def _extract_cv_text(file_path: str) -> str:
    ext = Path(file_path).suffix.lower()
    if ext == ".pdf":
        return _extract_pdf_text(file_path)
    if ext == ".docx":
        return _extract_docx_text(file_path)
    raise ValueError(f"Unsupported file type '{ext}'. Only PDF and DOCX are supported.")


# ---------------------------------------------------------------------------
# Skill 2 — LLM CV field extraction
# ---------------------------------------------------------------------------
class _CVFields(BaseModel):
    candidate_name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    latest_company: Optional[str] = None
    latest_position: Optional[str] = None


def _parse_cv_with_llm(text: str) -> _CVFields:
    structured_llm = llm.with_structured_output(_CVFields)
    return structured_llm.invoke(
        "Extract the following fields from the CV below. "
        "Return null for any field that cannot be clearly determined — do NOT guess.\n\n"
        f"CV:\n{text[:6000]}"
    )


# ---------------------------------------------------------------------------
# Skill 2 — Folder path parsing
# ---------------------------------------------------------------------------
def _parse_folder_path(file_path: str) -> dict:
    """Expects: .../[Dept] - [Position] - [JobCode]/[SubFolder]/filename"""
    parts = Path(os.path.normpath(file_path)).parts
    source = job_folder = None
    for i, part in enumerate(parts):
        if part in JOB_SUB_FOLDERS and i > 0:
            source = part
            job_folder = parts[i - 1]
            break
    if not source or not job_folder:
        return {}
    segments = job_folder.split(" - ", 2)
    if len(segments) < 3:
        return {}
    return {
        "processed_team": segments[0].strip(),
        "processed_position": segments[1].strip(),
        "request_code": segments[2].strip(),
        "source": source,
    }


# ---------------------------------------------------------------------------
# Skill 2 — Personal DB Excel operations
# ---------------------------------------------------------------------------
def _check_duplicate(email: str, request_code: str) -> dict | None:
    if not EXCEL_PATH or not os.path.exists(EXCEL_PATH):
        return None
    wb = _open_wb(EXCEL_PATH, read_only=True)
    if "Database" not in wb.sheetnames:
        wb.close()
        return None
    ws = wb["Database"]
    headers = None
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if headers is None:
            headers = [str(c).strip() if c is not None else "" for c in row]
            continue
        row_dict = dict(zip(headers, row))
        if (str(row_dict.get("Email") or "").strip().lower() == email.strip().lower()
                and str(row_dict.get("Request code") or "").strip() == request_code.strip()):
            wb.close()
            return {"row_number": row_idx, "candidate_name": row_dict.get("Candidate name")}
    wb.close()
    return None


def _add_row(row_data: dict) -> int:
    if not EXCEL_PATH:
        raise ValueError("EXCEL_PATH is not configured in .env")
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel file not found: {EXCEL_PATH}")
    wb = _open_wb(EXCEL_PATH)
    if "Database" not in wb.sheetnames:
        wb.create_sheet("Database")
    ws = wb["Database"]
    headers = _get_headers(ws)
    if not any(headers):
        for col, h in enumerate(DB_COLUMNS, 1):
            ws.cell(row=1, column=col, value=h)
        headers = DB_COLUMNS
    next_row = ws.max_row + 1
    if "No" in headers:
        ws.cell(row=next_row, column=headers.index("No") + 1, value=next_row - 1)
    for col_name, value in row_data.items():
        if col_name == "Profile":
            continue  # handled separately below
        if col_name in headers:
            ws.cell(row=next_row, column=headers.index(col_name) + 1, value=value)
    # Profile: clickable hyperlink — display text = candidate name, target = local file path
    if "Profile" in headers and row_data.get("Profile"):
        file_path_val = str(row_data["Profile"])
        display = str(row_data.get("Candidate name") or os.path.basename(file_path_val))
        url = "file:///" + file_path_val.replace("\\", "/")
        cell = ws.cell(row=next_row, column=headers.index("Profile") + 1, value=display)
        cell.hyperlink = url
        cell.font = Font(color="0563C1", underline="single")
    wb.save(EXCEL_PATH)
    wb.close()
    return next_row


def _overwrite_row(row_number: int, row_data: dict):
    wb = _open_wb(EXCEL_PATH)
    ws = wb["Database"]
    headers = _get_headers(ws)
    for col_name, value in row_data.items():
        if col_name in headers:
            ws.cell(row=row_number, column=headers.index(col_name) + 1, value=value)
    wb.save(EXCEL_PATH)
    wb.close()


# ---------------------------------------------------------------------------
# Skill 2 — Processed-files tracking
# ---------------------------------------------------------------------------
def _load_processed() -> set:
    if os.path.exists(PROCESSED_FILES_LOG):
        with open(PROCESSED_FILES_LOG) as f:
            return set(json.load(f))
    return set()


def _mark_processed(file_path: str):
    processed = _load_processed()
    processed.add(os.path.normpath(file_path))
    with open(PROCESSED_FILES_LOG, "w") as f:
        json.dump(list(processed), f, indent=2)


# ---------------------------------------------------------------------------
# Skill 2 — Core CV processing pipeline
# ---------------------------------------------------------------------------
def _build_row_data(cv: _CVFields, folder: dict, file_path: str) -> dict:
    return {
        "Request code": folder.get("request_code"),
        "Candidate name": cv.candidate_name,
        "Processed Team": folder.get("processed_team"),
        "Processed Position": folder.get("processed_position"),
        "Entry date": datetime.now().strftime("%Y-%m-%d"),
        "Source": folder.get("source"),
        "Referrer": None,  # left blank; recruiter fills if source = Referral
        "Latest company": cv.latest_company,
        "Latest position": cv.latest_position,
        "Email": cv.email,
        "Phone": cv.phone,
        "TA Incharge": TA_INCHARGE,
        "Profile": os.path.normpath(file_path),
    }


def _process_cv(file_path: str) -> dict:
    file_path = _resolve_path(file_path)
    result: dict = {"file": os.path.basename(file_path), "status": None, "messages": []}

    try:
        text = _extract_cv_text(file_path)
    except Exception as e:
        result.update(status="error", messages=[f"Cannot extract text: {e}"])
        _alerts.append({**result, "timestamp": datetime.now().isoformat()})
        return result

    if not text.strip():
        result.update(status="error", messages=["CV appears empty or image-only — could not extract text."])
        _alerts.append({**result, "timestamp": datetime.now().isoformat()})
        return result

    folder = _parse_folder_path(file_path)
    if not folder:
        result.update(status="error", messages=["Could not parse folder path. Expected: [Dept] - [Position] - [JobCode]/[SubFolder]/file"])
        _alerts.append({**result, "timestamp": datetime.now().isoformat()})
        return result

    try:
        cv = _parse_cv_with_llm(text)
    except Exception as e:
        result.update(status="error", messages=[f"LLM extraction failed: {e}"])
        _alerts.append({**result, "timestamp": datetime.now().isoformat()})
        return result

    missing = [label for label, val in [
        ("Candidate name", cv.candidate_name), ("Email", cv.email), ("Phone", cv.phone),
        ("Latest company", cv.latest_company), ("Latest position", cv.latest_position),
    ] if not val]
    if missing:
        result["messages"].append(f"⚠ Missing fields (left blank): {', '.join(missing)}")

    if cv.email:
        dup = _check_duplicate(cv.email, folder["request_code"])
        if dup:
            result.update(status="duplicate", messages=result["messages"] + [
                f"⚠ DUPLICATE: {cv.candidate_name or 'Unknown'} ({cv.email}) already exists "
                f"in row {dup['row_number']} for {folder['request_code']}. "
                "Use resolve_duplicate to keep or overwrite."
            ])
            _alerts.append({
                **result, "timestamp": datetime.now().isoformat(),
                "file_path": os.path.normpath(file_path),
                "cv_fields": cv.model_dump(), "folder_info": folder,
                "duplicate_row": dup["row_number"],
            })
            _mark_processed(file_path)
            return result

    row_data = _build_row_data(cv, folder, file_path)
    if EXCEL_PATH:
        try:
            new_row = _add_row(row_data)
            result.update(status="success")
            result["messages"].append(f"✓ Added {cv.candidate_name or 'candidate'} to database at row {new_row}.")
        except Exception as e:
            result.update(status="error")
            result["messages"].append(f"Excel write failed: {e}")
    else:
        result.update(status="no_excel")
        result["messages"].append("EXCEL_PATH not set — candidate NOT written to database.")

    _alerts.append({**result, "timestamp": datetime.now().isoformat()})
    _mark_processed(file_path)
    return result


# ---------------------------------------------------------------------------
# Skill 2 — Background CV watcher
# ---------------------------------------------------------------------------
def _scan_for_new_cvs() -> list[str]:
    processed = _load_processed()
    new_files = []
    if not os.path.isdir(JOBS_BASE_DIR):
        return []
    for job_folder in os.listdir(JOBS_BASE_DIR):
        job_path = os.path.join(JOBS_BASE_DIR, job_folder)
        if not os.path.isdir(job_path):
            continue
        for sub in JOB_SUB_FOLDERS:
            sub_path = os.path.join(job_path, sub)
            if not os.path.isdir(sub_path):
                continue
            for fname in os.listdir(sub_path):
                if fname.lower().endswith((".pdf", ".docx")):
                    fpath = os.path.normpath(os.path.join(sub_path, fname))
                    if fpath not in processed:
                        new_files.append(fpath)
    return new_files


async def _cv_watcher_loop():
    loop = asyncio.get_event_loop()
    while True:
        try:
            for fpath in _scan_for_new_cvs():
                await loop.run_in_executor(None, _process_cv, fpath)
        except Exception as e:
            _alerts.append({"status": "watcher_error", "messages": [str(e)], "timestamp": datetime.now().isoformat()})
        await asyncio.sleep(CV_POLL_INTERVAL)


# ---------------------------------------------------------------------------
# Skill 3 — Sync state tracking
# ---------------------------------------------------------------------------
def _row_key(row: dict) -> str:
    email = str(row.get("Email") or "").strip().lower()
    code = str(row.get("Request code") or "").strip()
    return f"{email}|{code}"


def _load_sync_state() -> dict:
    if os.path.exists(SYNC_STATE_FILE):
        with open(SYNC_STATE_FILE) as f:
            return json.load(f)
    return {"last_sync": None, "rows": {}}


def _save_sync_state(personal_rows: list[dict]):
    state = {
        "last_sync": datetime.now().isoformat(),
        "rows": {
            _row_key(r): {k: str(v) if v is not None else None for k, v in r.items()}
            for r in personal_rows
            if _row_key(r) not in ("|", "")
        },
    }
    with open(SYNC_STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)


def _row_changed_vs_snapshot(current: dict, snapshot: dict) -> bool:
    for col in SYNC_COLUMNS:
        if str(current.get(col) or "").strip() != str(snapshot.get(col) or "").strip():
            return True
    return False


# ---------------------------------------------------------------------------
# Skill 3 — Personal & team DB readers
# ---------------------------------------------------------------------------
def _read_personal_db() -> list[dict]:
    if not EXCEL_PATH or not os.path.exists(EXCEL_PATH):
        return []
    wb = _open_wb(EXCEL_PATH, read_only=True)
    if "Database" not in wb.sheetnames:
        wb.close()
        return []
    ws = wb["Database"]
    headers = None
    rows = []
    for row in ws.iter_rows(values_only=True):
        if headers is None:
            headers = [str(c).strip() if c is not None else "" for c in row]
            continue
        if not any(c is not None for c in row):
            continue
        rows.append(dict(zip(headers, row)))
    wb.close()
    return rows


def _read_team_db() -> dict[str, tuple[int, dict]]:
    """Returns {row_key: (excel_row_number, row_dict)} for every data row in team DB."""
    if not TEAM_EXCEL_PATH or not os.path.exists(TEAM_EXCEL_PATH):
        return {}
    wb = _open_wb(TEAM_EXCEL_PATH, read_only=True)
    if "Database" not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb["Database"]
    lookup: dict[str, tuple[int, dict]] = {}
    headers = None
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if headers is None:
            headers = [str(c).strip() if c is not None else "" for c in row]
            continue
        if not any(c is not None for c in row):
            continue
        row_dict = dict(zip(headers, row))
        key = _row_key(row_dict)
        if key and key != "|":
            lookup[key] = (row_idx, row_dict)
    wb.close()
    return lookup


# ---------------------------------------------------------------------------
# Skill 3 — Team DB write operations
# ---------------------------------------------------------------------------
def _add_team_row(row_data: dict) -> int:
    if not TEAM_EXCEL_PATH or not os.path.exists(TEAM_EXCEL_PATH):
        raise FileNotFoundError(f"Team Excel not found: {TEAM_EXCEL_PATH}")
    wb = _open_wb(TEAM_EXCEL_PATH)
    if "Database" not in wb.sheetnames:
        wb.create_sheet("Database")
    ws = wb["Database"]
    headers = _get_headers(ws)
    if not any(headers):
        for col, h in enumerate(DB_COLUMNS, 1):
            ws.cell(row=1, column=col, value=h)
        headers = DB_COLUMNS
    next_row = ws.max_row + 1
    if "No" in headers:
        ws.cell(row=next_row, column=headers.index("No") + 1, value=next_row - 1)
    for col_name in SYNC_COLUMNS:
        value = row_data.get(col_name)
        if col_name in headers and value is not None:
            ws.cell(row=next_row, column=headers.index(col_name) + 1, value=value)
    wb.save(TEAM_EXCEL_PATH)
    wb.close()
    return next_row


def _update_team_row(row_number: int, changed_fields: dict):
    """Write only the changed fields into the existing team DB row."""
    wb = _open_wb(TEAM_EXCEL_PATH)
    ws = wb["Database"]
    headers = _get_headers(ws)
    for field, vals in changed_fields.items():
        if field in headers:
            ws.cell(row=row_number, column=headers.index(field) + 1, value=vals["mine"])
    wb.save(TEAM_EXCEL_PATH)
    wb.close()


def _write_conflict_log(conflicts: list[dict]):
    """Append conflict rows to the Conflict Log sheet with red highlighting."""
    wb = _open_wb(TEAM_EXCEL_PATH)
    if CONFLICT_SHEET not in wb.sheetnames:
        ws_conflict = wb.create_sheet(CONFLICT_SHEET)
        for col, h in enumerate(CONFLICT_LOG_COLUMNS, 1):
            cell = ws_conflict.cell(row=1, column=col, value=h)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(fill_type="solid", fgColor="FFD700")
    else:
        ws_conflict = wb[CONFLICT_SHEET]

    red = PatternFill(fill_type="solid", fgColor="FF6B6B")
    next_row = ws_conflict.max_row + 1
    for c in conflicts:
        values = [
            c["timestamp"], c["email"], c["request_code"], c["candidate_name"],
            c["field"], c["my_value"], c["team_value"], c["my_ta"], c["team_ta"],
        ]
        for col, val in enumerate(values, 1):
            cell = ws_conflict.cell(row=next_row, column=col, value=val)
            cell.fill = red
        # Also add a comment on the Email cell in Database sheet to surface the conflict inline
        if "Database" in wb.sheetnames:
            ws_db = wb["Database"]
            db_headers = _get_headers(ws_db)
            # Find the conflicted row in Database by scanning (best effort)
            if "Email" in db_headers:
                email_col = db_headers.index("Email") + 1
                for db_row in range(2, ws_db.max_row + 1):
                    cell_val = ws_db.cell(row=db_row, column=email_col).value
                    if str(cell_val or "").strip().lower() == c["email"].lower():
                        target = ws_db.cell(row=db_row, column=email_col)
                        note = f"Sync conflict {c['timestamp'][:10]}: field '{c['field']}' differs (mine: {c['my_value']} | team: {c['team_value']}). Check Conflict Log sheet."
                        target.comment = Comment(note, "Recruiter Agent")
                        break
        next_row += 1

    wb.save(TEAM_EXCEL_PATH)
    wb.close()


# ---------------------------------------------------------------------------
# Skill 3 — Sync logic
# ---------------------------------------------------------------------------
def _get_changed_fields(my_row: dict, team_row: dict) -> dict:
    """Return {field: {mine: ..., theirs: ...}} for every field that differs."""
    return {
        col: {"mine": str(my_row.get(col) or "").strip(), "theirs": str(team_row.get(col) or "").strip()}
        for col in SYNC_COLUMNS
        if str(my_row.get(col) or "").strip() != str(team_row.get(col) or "").strip()
    }


def _is_conflict(team_row: dict, today: str) -> bool:
    """Conflict: team DB row was added today by a different TA (two recruiters processed same candidate today)."""
    team_ta = str(team_row.get("TA Incharge") or "").strip().lower()
    team_date = str(team_row.get("Entry date") or "").strip()[:10]
    return team_date == today and team_ta != (TA_INCHARGE or "").lower().strip()


def _validate_excel(path: str) -> bool:
    """Return True if the file can be opened as a valid Excel workbook."""
    try:
        wb = _open_wb(path, read_only=True)
        wb.close()
        return True
    except Exception:
        return False


def _run_sync() -> dict:
    today = datetime.now().strftime("%Y-%m-%d")
    result = {"date": today, "timestamp": datetime.now().isoformat(),
               "added": 0, "updated": 0, "skipped": 0, "conflicts": 0, "errors": []}

    if not TEAM_EXCEL_PATH or not os.path.exists(TEAM_EXCEL_PATH):
        result["errors"].append(f"TEAM_EXCEL_PATH not found: {TEAM_EXCEL_PATH}")
        return result
    if not EXCEL_PATH or not os.path.exists(EXCEL_PATH):
        result["errors"].append(f"Personal EXCEL_PATH not found: {EXCEL_PATH}")
        return result
    if not _validate_excel(TEAM_EXCEL_PATH):
        result["errors"].append(
            f"Team Excel file is not a valid workbook: {TEAM_EXCEL_PATH}. "
            "Open the file in Excel, add a sheet named 'Database', and save it once to make it a valid .xlsm file."
        )
        return result
    if not _validate_excel(EXCEL_PATH):
        result["errors"].append(f"Personal Excel file is not a valid workbook: {EXCEL_PATH}.")
        return result

    personal_rows = _read_personal_db()
    team_lookup = _read_team_db()
    snapshot = _load_sync_state().get("rows", {})

    # Determine which personal rows to sync: added today OR changed since last sync
    to_sync = []
    for row in personal_rows:
        entry_date = str(row.get("Entry date") or "").strip()[:10]
        key = _row_key(row)
        if not key or key == "|":
            continue
        if entry_date == today:
            to_sync.append(row)
        elif key in snapshot and _row_changed_vs_snapshot(row, snapshot[key]):
            to_sync.append(row)

    conflicts_to_log: list[dict] = []

    for my_row in to_sync:
        key = _row_key(my_row)
        email = str(my_row.get("Email") or "").strip().lower()
        request_code = str(my_row.get("Request code") or "").strip()
        candidate_name = str(my_row.get("Candidate name") or "")
        try:
            if key not in team_lookup:
                _add_team_row(my_row)
                # Keep lookup fresh so duplicate my_rows in the same batch don't double-insert
                team_lookup[key] = (-1, my_row)
                result["added"] += 1
            else:
                team_row_num, team_row = team_lookup[key]
                changed = _get_changed_fields(my_row, team_row)
                if not changed:
                    result["skipped"] += 1
                    continue
                if _is_conflict(team_row, today):
                    for field, vals in changed.items():
                        conflicts_to_log.append({
                            "timestamp": datetime.now().isoformat(),
                            "email": email,
                            "request_code": request_code,
                            "candidate_name": candidate_name,
                            "field": field,
                            "my_value": vals["mine"],
                            "team_value": vals["theirs"],
                            "my_ta": TA_INCHARGE,
                            "team_ta": str(team_row.get("TA Incharge") or ""),
                        })
                    result["conflicts"] += 1
                else:
                    _update_team_row(team_row_num, changed)
                    result["updated"] += 1
        except Exception as e:
            result["errors"].append(f"{candidate_name} ({email}): {e}")

    if conflicts_to_log:
        try:
            _write_conflict_log(conflicts_to_log)
        except Exception as e:
            result["errors"].append(f"Conflict log write failed: {e}")

    _save_sync_state(personal_rows)

    summary = (f"Daily sync {today}: {result['added']} added, {result['updated']} updated, "
               f"{result['skipped']} skipped, {result['conflicts']} conflicts.")
    _alerts.append({"status": "sync_complete", "messages": [summary],
                    "timestamp": result["timestamp"], "details": result})
    return result


# ---------------------------------------------------------------------------
# Skill 3 — Daily sync scheduler
# ---------------------------------------------------------------------------
def _get_scheduled_time() -> tuple[int, int] | None:
    return SYNC_SCHEDULE.get((TA_INCHARGE or "").lower().strip())


async def _daily_sync_scheduler():
    loop = asyncio.get_event_loop()
    last_sync_date = None
    while True:
        try:
            scheduled = _get_scheduled_time()
            if scheduled:
                now = datetime.now()
                h, m = scheduled
                if (now.hour, now.minute) >= (h, m) and now.date() != last_sync_date:
                    await loop.run_in_executor(None, _run_sync)
                    last_sync_date = now.date()
        except Exception as e:
            _alerts.append({"status": "sync_error", "messages": [f"Scheduler error: {e}"],
                            "timestamp": datetime.now().isoformat()})
        await asyncio.sleep(30)  # check every 30 s


# ---------------------------------------------------------------------------
# Lifespan — start both background tasks
# ---------------------------------------------------------------------------
@contextlib.asynccontextmanager
async def _lifespan(app_instance):
    cv_task = asyncio.create_task(_cv_watcher_loop())
    sync_task = asyncio.create_task(_daily_sync_scheduler())
    yield
    cv_task.cancel()
    sync_task.cancel()
    for t in (cv_task, sync_task):
        try:
            await t
        except asyncio.CancelledError:
            pass


# ---------------------------------------------------------------------------
# App + LLM
# ---------------------------------------------------------------------------
app = GreenNodeAgentBaseApp(lifespan=_lifespan)

LLM_MODEL = os.environ.get("LLM_MODEL", "")
LLM_BASE_URL = os.environ.get("LLM_BASE_URL", "")
LLM_API_KEY = os.environ.get("LLM_API_KEY", "")
if not LLM_MODEL or not LLM_BASE_URL or not LLM_API_KEY:
    raise ValueError(
        "LLM_MODEL, LLM_BASE_URL, and LLM_API_KEY are required. "
        "Set them in .env or use /agentbase-llm to get a platform API key."
    )

llm = ChatOpenAI(model=LLM_MODEL, base_url=LLM_BASE_URL, api_key=LLM_API_KEY)


# ---------------------------------------------------------------------------
# LangGraph tools — Skill 1
# ---------------------------------------------------------------------------
@tool
def create_job_folder(folder_name: str) -> str:
    """Create a job folder with standard recruitment sub-folders: LinkedIn, VNG Careers, Referral, TA Search, Others.

    Args:
        folder_name: Name of the job folder (e.g. 'ZDA - Data Scientist - 26-ZDA-3117').
    """
    job_path = os.path.join(JOBS_BASE_DIR, folder_name)
    if os.path.exists(job_path):
        return f"Folder '{folder_name}' already exists."
    os.makedirs(job_path)
    for sub in JOB_SUB_FOLDERS:
        os.makedirs(os.path.join(job_path, sub))
    return f"Created '{folder_name}' with sub-folders: {', '.join(JOB_SUB_FOLDERS)}."


# ---------------------------------------------------------------------------
# LangGraph tools — Skill 2
# ---------------------------------------------------------------------------
@tool
def process_cv_file(file_path: str) -> str:
    """Manually process a CV file (PDF or DOCX) and add the candidate to the personal recruitment database.

    Args:
        file_path: Full path to the CV file.
    """
    result = _process_cv(file_path)
    return "\n".join([f"Status: {result['status']}"] + result["messages"])


@tool
def get_alerts() -> str:
    """Get all pending alerts: CV processing results, duplicates, sync outcomes, errors."""
    if not _alerts:
        return "No pending alerts."
    lines = []
    for i, alert in enumerate(_alerts[-20:], 1):
        lines.append(f"[{i}] {alert.get('timestamp', '')[:19]} | {alert.get('status', '')} | {alert.get('file', 'system')}")
        for msg in alert.get("messages", []):
            lines.append(f"    {msg}")
    return "\n".join(lines)


@tool
def clear_alerts() -> str:
    """Clear all pending alerts."""
    _alerts.clear()
    return "All alerts cleared."


@tool
def resolve_duplicate(email: str, request_code: str, action: str) -> str:
    """Resolve a duplicate candidate detected during CV processing.

    Args:
        email: The candidate's email address.
        request_code: The job request code (e.g. '26-ZDA-3117').
        action: 'keep' to keep the existing row, or 'overwrite' to replace with new CV data.
    """
    pending = next(
        (a for a in _alerts
         if a.get("status") == "duplicate"
         and str(a.get("cv_fields", {}).get("email") or "").lower() == email.strip().lower()
         and str(a.get("folder_info", {}).get("request_code") or "") == request_code.strip()),
        None,
    )
    if not pending:
        return f"No pending duplicate for '{email}' / '{request_code}'."
    name = pending["cv_fields"].get("candidate_name") or "Unknown"
    row_num = pending["duplicate_row"]
    if action.lower() == "keep":
        _alerts.remove(pending)
        return f"Kept existing row {row_num} for {name}. No changes made."
    if action.lower() == "overwrite":
        row_data = _build_row_data(_CVFields(**pending["cv_fields"]), pending["folder_info"], pending["file_path"])
        try:
            _overwrite_row(row_num, row_data)
        except Exception as e:
            return f"Overwrite failed: {e}"
        _alerts.remove(pending)
        return f"Overwrote row {row_num} with updated data for {name}."
    return "Invalid action. Use 'keep' or 'overwrite'."


# ---------------------------------------------------------------------------
# LangGraph tools — Skill 3
# ---------------------------------------------------------------------------
@tool
def run_sync_now() -> str:
    """Manually trigger the daily sync from personal database to team database right now."""
    result = _run_sync()
    lines = [f"Sync completed at {result['timestamp'][:19]}",
             f"  Added:     {result['added']}",
             f"  Updated:   {result['updated']}",
             f"  Skipped:   {result['skipped']} (no changes)",
             f"  Conflicts: {result['conflicts']} (written to '{CONFLICT_SHEET}' sheet in team DB)"]
    if result["errors"]:
        lines.append(f"  Errors: {'; '.join(result['errors'])}")
    return "\n".join(lines)


@tool
def get_sync_status() -> str:
    """Show when the last sync ran and when the next one is scheduled for this TA."""
    state = _load_sync_state()
    last = (state.get("last_sync") or "Never")[:19]
    scheduled = _get_scheduled_time()
    sched_str = f"{scheduled[0]:02d}:{scheduled[1]:02d} daily" if scheduled else f"Not scheduled (TA '{TA_INCHARGE}' not in schedule)"
    rows_tracked = len(state.get("rows", {}))
    return (f"TA Incharge:    {TA_INCHARGE}\n"
            f"Scheduled sync: {sched_str}\n"
            f"Last sync:      {last}\n"
            f"Rows tracked:   {rows_tracked}")


# ---------------------------------------------------------------------------
# LangGraph graph
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """You are a Recruiter Assistant Agent for VNG recruitment operations.

Skills:
1. **create_job_folder** — create job folder with standard sub-folders (LinkedIn, VNG Careers, Referral, TA Search, Others)
2. **process_cv_file** — process a CV file and add candidate to personal database
3. **get_alerts** — check pending alerts: CV results, duplicates, sync outcomes, errors
4. **clear_alerts** — dismiss resolved alerts
5. **resolve_duplicate** — resolve a duplicate candidate: keep existing row or overwrite
6. **run_sync_now** — manually trigger personal → team database sync immediately
7. **get_sync_status** — show last sync time and next scheduled sync

Rules:
- Never fill Stage, Status, Note, Reason for failure, or salary fields — recruiter only
- Never delete existing rows; only add or (on explicit recruiter instruction) overwrite
- Missing CV fields: leave blank, alert recruiter — do not guess
- Referrer: blank unless source = Referral
- Sync: only Excel data, never copy CV files"""


class State(TypedDict):
    messages: Annotated[list, add_messages]


tools = [
    create_job_folder,
    process_cv_file, get_alerts, clear_alerts, resolve_duplicate,
    run_sync_now, get_sync_status,
]
llm_with_tools = llm.bind_tools(tools)


def chatbot(state: State) -> dict:
    messages = [SystemMessage(content=SYSTEM_PROMPT)] + state["messages"]
    return {"messages": [llm_with_tools.invoke(messages)]}


graph_builder = StateGraph(State)
graph_builder.add_node("chatbot", chatbot)
graph_builder.add_node("tools", ToolNode(tools))
graph_builder.add_edge(START, "chatbot")
graph_builder.add_conditional_edges("chatbot", tools_condition)
graph_builder.add_edge("tools", "chatbot")
graph = graph_builder.compile()


# ---------------------------------------------------------------------------
# Entrypoint & health check
# ---------------------------------------------------------------------------
@app.entrypoint
def handler(payload: dict, context: RequestContext) -> dict:
    message = payload.get("message", "Hello")
    result = graph.invoke({"messages": [("user", message)]})
    return {
        "status": "success",
        "response": result["messages"][-1].content,
        "timestamp": datetime.now().isoformat(),
    }


@app.ping
def health_check() -> PingStatus:
    return PingStatus.HEALTHY


if __name__ == "__main__":
    app.run(port=8080, host="0.0.0.0")

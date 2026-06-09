# =============================================================================
# PUBLIC RELEASE SAFETY NOTE
# =============================================================================
# This public version intentionally uses a mocked work-item creation service.
# Replace WorkItemService with your own approved ticket/work-item API integration.
#
# Do not commit runtime logs, local SQLite databases, real workbook links,
# internal field names, exported spreadsheet data, or any company-specific
# system names. Add generated files such as *.log, *.db, *.sqlite, and
# *.sqlite3 to .gitignore before publishing this script.
# =============================================================================

# =============================================================================
# INSTALL / RUNTIME DEPENDENCIES
# =============================================================================
# Standard library:
# - os, re, sys, time, queue, sqlite3, hashlib, logging, threading, urllib.parse
# - dataclasses, datetime, typing
#
# Built-in GUI modules:
# - tkinter
# - tkinter.ttk
# - tkinter.scrolledtext
#
# External packages:
#   pip install pywin32
#
# Local runtime requirements:
# - Microsoft Excel installed on Windows
#
# Notes:
# - This is a Windows-only desktop app due to Excel COM automation.
# - The downstream work-item creation is currently mocked intentionally.

import os
import re
import sys
import time
import queue
import sqlite3
import hashlib
import logging
import threading
import urllib.parse
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pythoncom
import win32com.client as win32

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.scrolledtext as scrolledtext


# =============================================================================
# CONFIGURATION
# =============================================================================

APP_TITLE = "Spreadsheet Work Item Automation - Flow UI"
LOG_FILE = "work_item_automation.log"
SQLITE_DB = "work_item_duplicates.db"

COMPLETED_STATUSES = ["completed", "complete", "done", "finished", "closed"]
HIGH_CONFIDENCE_THRESHOLD = 0.70
DEFAULT_WORK_ITEM_REF_COLUMN = "Work Item Reference/ID"

WORK_ITEM_FIELDS = {
    "status": "Task Status",
    "work_item_ref": "Work Item Reference/ID",
    "owner": "Task Owner/Assignee",
    "station_name": "System/Station/Machine Name",
    "description": "Task Description/Details",
    "completed_date": "Task Completion Date",
    "priority": "Task Priority",
}

AUTO_MAPPING_PATTERNS = {
    "status": [r"\bstatus\b", r"\bstate\b", r"\bcondition\b", r"\bprogress\b", r"\bstage\b"],
    "work_item_ref": [r"\bticket\b", r"\breference\b", r"\bref\b", r"\btracking\b", r"\bnumber\b", r"\bwork item\b"],
    "owner": [r"\bowner\b", r"\bassignee\b", r"\bassigned\b", r"\bresponsible\b", r"\blead\b", r"\bcontact\b"],
    "station_name": [r"\bstation\b", r"\bsystem\b", r"\bmachine\b", r"\bhost\b", r"\bdevice\b", r"\bequipment\b", r"\basset\b", r"\bserver\b"],
    "description": [r"\bdescription\b", r"\bdesc\b", r"\bdetails\b", r"\bsummary\b", r"\btitle\b", r"\bissue\b", r"\bproblem\b"],
    "completed_date": [r"\bcompleted\b", r"\bcompletion\b", r"\bdone\b", r"\bfinished\b", r"\bclosed\b", r"\bdate\b", r"\btimestamp\b"],
    "priority": [r"\bpriority\b", r"\bprio\b", r"\bseverity\b", r"\burgency\b", r"\bimpact\b"],
}


# =============================================================================
# LOGGING
# =============================================================================

def setup_logging() -> logging.Logger:
    """Create the shared application logger."""
    logger = logging.getLogger("work_item_automation")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setFormatter(formatter)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(formatter)
    logger.addHandler(fh)
    logger.addHandler(sh)
    return logger


logger = setup_logging()


# =============================================================================
# DATA MODELS
# =============================================================================

@dataclass
class ColumnMapping:
    """Resolved mapping from logical business fields to Excel headers."""
    status: Optional[str] = None
    work_item_ref: Optional[str] = None
    owner: Optional[str] = None
    station_name: Optional[str] = None
    description: Optional[str] = None
    completed_date: Optional[str] = None
    priority: Optional[str] = None
    confidence: float = 0.0


@dataclass
class TaskProposal:
    """Single row projected into a proposed work item."""
    proposal_id: str
    row_index: int
    excel_row: int
    status: str
    owner: str
    station_name: str
    description: str
    completed_date: str
    priority: str
    confidence: float
    source_data: Dict[str, Any]
    has_existing_work_item: bool = False
    existing_work_item_id: str = ""
    recommended_action: str = "skip"
    task_hash: str = ""
    station_code: str = ""
    decision_reason: str = ""


@dataclass
class WorkbookRef:
    """Workbook identity stored without long-lived COM proxies."""
    workbook_name: str = ""
    workbook_full_name: str = ""
    workbook_path: str = ""
    is_sharepoint: bool = False
    source: str = ""
    worksheet_names: List[str] = field(default_factory=list)
    worksheet_hint: str = ""
    detected_active_sheet: str = ""


@dataclass
class AppConfig:
    """Runtime behavior toggles."""
    create_only_for_completed: bool = False
    min_description_length: int = 3
    auto_create_work_item_ref_column: bool = True
    work_item_ref_column_name: str = DEFAULT_WORK_ITEM_REF_COLUMN
    confirm_before_open_excel: bool = True
    confirm_before_create_work_items: bool = True
    confirm_before_writeback: bool = True


# =============================================================================
# DUPLICATE PROTECTION
# =============================================================================

class DuplicateProtection:
    """Local duplicate protection using SQLite."""
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        """Create schema and apply additive migrations."""
        conn = sqlite3.connect(self.db_path)
        try:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS processed_tasks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_hash TEXT UNIQUE NOT NULL,
                    work_item_ref TEXT NOT NULL,
                    created_date TEXT NOT NULL,
                    owner TEXT,
                    station_name TEXT,
                    completed_date TEXT,
                    description_preview TEXT
                )
            """)
            conn.commit()
            self._migrate_schema(conn)
        finally:
            conn.close()

    def _migrate_schema(self, conn):
        """Add newly introduced columns to older local DB files."""
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(processed_tasks)")
        existing_columns = {row[1] for row in cur.fetchall()}
        expected_columns = {
            "owner": "TEXT",
            "station_name": "TEXT",
            "completed_date": "TEXT",
            "description_preview": "TEXT",
        }
        for col_name, col_type in expected_columns.items():
            if col_name not in existing_columns:
                logger.info(f"🔧 Migrating SQLite schema: adding column '{col_name}'")
                cur.execute(f"ALTER TABLE processed_tasks ADD COLUMN {col_name} {col_type}")
        conn.commit()

    @staticmethod
    def normalize_text(value: Any) -> str:
        """Normalize free text before hashing."""
        text = str(value or "").strip().lower()
        text = re.sub(r"\s+", " ", text)
        return text

    def compute_task_hash(self, task_data: Dict[str, Any]) -> str:
        """Build a stable duplicate-protection hash from task-defining fields."""
        fingerprint = "|".join([
            self.normalize_text(task_data.get("owner")),
            self.normalize_text(task_data.get("station_name")),
            self.normalize_text(task_data.get("description")),
            self.normalize_text(task_data.get("completed_date")),
            self.normalize_text(task_data.get("priority")),
        ])
        return hashlib.sha256(fingerprint.encode("utf-8")).hexdigest()[:20]

    def is_duplicate(self, task_hash: str) -> Tuple[bool, Optional[str]]:
        """Return whether the task hash already exists and its work-item reference."""
        conn = sqlite3.connect(self.db_path)
        try:
            cur = conn.cursor()
            cur.execute("SELECT work_item_ref FROM processed_tasks WHERE task_hash = ?", (task_hash,))
            result = cur.fetchone()
            return (True, result[0]) if result else (False, None)
        finally:
            conn.close()

    def record_processed_task(self, task_hash: str, work_item_ref: str, task_data: Dict[str, Any]):
        """Persist a successful creation outcome for future duplicate prevention."""
        conn = sqlite3.connect(self.db_path)
        try:
            cur = conn.cursor()
            cur.execute("""
                INSERT OR REPLACE INTO processed_tasks
                (task_hash, work_item_ref, created_date, owner, station_name, completed_date, description_preview)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                task_hash,
                work_item_ref,
                datetime.now().isoformat(),
                str(task_data.get("owner", ""))[:100],
                str(task_data.get("station_name", ""))[:100],
                str(task_data.get("completed_date", ""))[:50],
                str(task_data.get("description", ""))[:200],
            ))
            conn.commit()
        finally:
            conn.close()


# =============================================================================
# COLUMN DETECTOR
# =============================================================================

class IntelligentColumnDetector:
    """Heuristic mapper from headers/content to logical task fields."""

    def detect_columns(self, headers: List[str], sample_data: List[List[Any]]) -> ColumnMapping:
        mapping = ColumnMapping()
        scores: Dict[str, Dict[str, float]] = {}
        normalized_headers = [str(h).strip() if h is not None else "" for h in headers]

        for i, header in enumerate(normalized_headers):
            header_lower = header.lower()
            col_scores: Dict[str, float] = {}

            for field_type, patterns in AUTO_MAPPING_PATTERNS.items():
                score = 0.0
                for pattern in patterns:
                    if re.search(pattern, header_lower):
                        score += 4 if re.fullmatch(pattern, header_lower) else 2
                col_scores[field_type] = score

            sample_values = []
            for row in sample_data[:10]:
                if i < len(row) and row[i] is not None:
                    sample_values.append(str(row[i]).strip().lower())

            if sample_values:
                if any(any(s in v for s in ["complete", "done", "closed", "pending", "in progress", "open"]) for v in sample_values):
                    col_scores["status"] = col_scores.get("status", 0) + 3
                if any(any(p in v for p in ["p1", "p2", "p3", "high", "medium", "low", "critical"]) for v in sample_values):
                    col_scores["priority"] = col_scores.get("priority", 0) + 3
                avg_len = sum(len(v) for v in sample_values) / max(1, len(sample_values))
                if avg_len > 20:
                    col_scores["description"] = col_scores.get("description", 0) + 2

            scores[header] = col_scores

        assigned_headers = set()
        field_assignments = {}
        priority_order = ["work_item_ref", "status", "priority", "owner", "station_name", "description", "completed_date"]

        for field_type in priority_order:
            best_header = None
            best_score = 0.0
            for header, col_scores in scores.items():
                if header in assigned_headers:
                    continue
                score = col_scores.get(field_type, 0)
                if score > best_score:
                    best_score = score
                    best_header = header
            if best_header and best_score > 0:
                setattr(mapping, field_type, best_header)
                field_assignments[field_type] = best_header
                assigned_headers.add(best_header)

        base_conf = len(field_assignments) / len(AUTO_MAPPING_PATTERNS) if AUTO_MAPPING_PATTERNS else 0.0
        mapping.confidence = min(1.0, base_conf + 0.05)
        return mapping


# =============================================================================
# TASK ANALYZER
# =============================================================================

class TaskAnalyzer:
    """Convert worksheet rows into editable, reviewable work-item proposals."""

    def __init__(self, duplicate_protection: DuplicateProtection, config: AppConfig):
        self.duplicate_protection = duplicate_protection
        self.config = config

    @staticmethod
    def extract_station_code(station_name: str) -> str:
        """Extract the first 3-digit code from station text if present."""
        match = re.search(r"(\d{3})", station_name or "")
        return match.group(1) if match else ""

    def analyze_row(self, row_data: List[Any], mapping: ColumnMapping, headers: List[str], row_index: int) -> TaskProposal:
        task_data = {}

        for field_name in WORK_ITEM_FIELDS.keys():
            col_name = getattr(mapping, field_name, None)
            if col_name and col_name in headers:
                idx = headers.index(col_name)
                task_data[field_name] = row_data[idx] if idx < len(row_data) else None
            else:
                task_data[field_name] = None

        status = str(task_data.get("status", "") or "").strip()
        owner = str(task_data.get("owner", "") or "").strip()
        station_name = str(task_data.get("station_name", "") or "").strip()
        description = str(task_data.get("description", "") or "").strip()
        completed_date = str(task_data.get("completed_date", "") or "").strip()
        priority = str(task_data.get("priority", "") or "").strip()
        existing_work_item = str(task_data.get("work_item_ref", "") or "").strip()
        station_code = self.extract_station_code(station_name)

        has_existing = bool(existing_work_item and existing_work_item.lower() not in ["none", "null", "", "n/a", "tbd"])
        task_hash = self.duplicate_protection.compute_task_hash(task_data)
        is_dup, dup_id = self.duplicate_protection.is_duplicate(task_hash)

        recommendation = "create"
        reason_parts = []

        if has_existing:
            recommendation = "skip"
            reason_parts.append("existing work item ref found")
        elif is_dup:
            recommendation = "skip"
            reason_parts.append("duplicate task hash found")
        else:
            if not station_name:
                reason_parts.append("missing station")
            if not description or len(description) < self.config.min_description_length:
                reason_parts.append("short/missing description")
            if self.config.create_only_for_completed:
                if not any(s in status.lower() for s in COMPLETED_STATUSES):
                    recommendation = "skip"
                    reason_parts.append("not completed while completed-only mode is enabled")
            else:
                if not status:
                    reason_parts.append("missing status allowed")
                elif "open" in status.lower():
                    reason_parts.append("open status allowed")

        confidence = self._calculate_confidence(task_data, mapping)

        if recommendation != "skip":
            has_core_signal = bool(station_name or description or owner or priority or status)
            if not has_core_signal:
                recommendation = "skip"
                reason_parts.append("no meaningful task content")

        proposal_id = hashlib.md5(f"{row_index}|{task_hash}".encode("utf-8")).hexdigest()[:12]

        return TaskProposal(
            proposal_id=proposal_id,
            row_index=row_index,
            excel_row=row_index + 2,
            status=status,
            owner=owner,
            station_name=station_name,
            description=description,
            completed_date=completed_date,
            priority=priority,
            confidence=confidence,
            source_data=task_data,
            has_existing_work_item=has_existing,
            existing_work_item_id=existing_work_item if has_existing else (dup_id or ""),
            recommended_action=recommendation,
            task_hash=task_hash,
            station_code=station_code,
            decision_reason="; ".join(reason_parts) if reason_parts else "eligible for creation"
        )

    def _calculate_confidence(self, task_data: Dict[str, Any], mapping: ColumnMapping) -> float:
        """Confidence is advisory; human review remains the decision point."""
        score = 0.0
        weights = {
            "owner": 0.10,
            "station_name": 0.25,
            "description": 0.30,
            "status": 0.10,
            "priority": 0.10,
            "completed_date": 0.05,
        }
        for field, weight in weights.items():
            value = str(task_data.get(field, "") or "").strip()
            if value:
                score += weight
        score += mapping.confidence * 0.20
        return min(1.0, score)


# =============================================================================
# WORK ITEM SERVICE (MOCK)
# =============================================================================

class WorkItemService:
    """Mock downstream work-item service used for full workflow testing."""

    def build_work_item_payload(self, proposal: TaskProposal) -> Dict[str, Any]:
        title_parts = []
        if proposal.station_code:
            title_parts.append(f"[{proposal.station_code}]")
        if proposal.priority:
            title_parts.append(f"[{proposal.priority}]")
        title_parts.append("Task")
        if proposal.description:
            title_parts.append(f"- {proposal.description[:80]}")
        title = " ".join(title_parts)

        description = (
            f"Station: {proposal.station_name or 'N/A'}\n"
            f"Station Code: {proposal.station_code or 'N/A'}\n"
            f"Owner: {proposal.owner or 'N/A'}\n"
            f"Priority: {proposal.priority or 'N/A'}\n"
            f"Status: {proposal.status or 'N/A'}\n"
            f"Completed Date: {proposal.completed_date or 'N/A'}\n\n"
            f"Original Description:\n{proposal.description or 'N/A'}"
        )

        return {
            "title": title,
            "description": description,
            "owner": proposal.owner,
            "priority": proposal.priority,
            "station_name": proposal.station_name,
            "station_code": proposal.station_code,
        }

    def create_work_item(self, proposal: TaskProposal) -> Tuple[str, Dict[str, Any]]:
        """Return a fake work-item ID plus the payload used to create it."""
        payload = self.build_work_item_payload(proposal)
        time.sleep(0.12)
        ts = int(time.time())
        station = proposal.station_code or "SYS"
        work_item_id = f"WORK-{ts}-{station}"
        return work_item_id, payload


# =============================================================================
# EXCEL COM SERVICE
# =============================================================================

class ExcelService:
    """Excel COM wrapper with thread-safe reacquisition of workbook/sheet objects."""

    def __init__(self, ui_callback=None):
        self.workbook_ref: Optional[WorkbookRef] = None
        self.current_worksheet_name: str = ""
        self.headers: List[str] = []
        self.data_rows: List[List[Any]] = []
        self.ui_callback = ui_callback

    def _log(self, msg: str):
        logger.info(msg)
        if self.ui_callback:
            self.ui_callback("log", msg)

    @staticmethod
    def _coinit():
        """Initialize COM for the current worker thread."""
        pythoncom.CoInitialize()

    @staticmethod
    def _couninit():
        """Release COM for the current worker thread."""
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    @staticmethod
    def _get_excel_app():
        """Attach to a running Excel instance or start one if needed."""
        try:
            return win32.GetActiveObject("Excel.Application")
        except Exception:
            app = win32.Dispatch("Excel.Application")
            app.Visible = True
            return app

    @staticmethod
    def _find_workbook(excel_app, workbook_name="", workbook_full_name=""):
        """Find workbook by full path or name inside current Excel process."""
        for i in range(1, excel_app.Workbooks.Count + 1):
            wb = excel_app.Workbooks(i)
            try:
                full_name = str(getattr(wb, "FullName", "")).lower()
                name = str(getattr(wb, "Name", "")).lower()
                if workbook_full_name and full_name == workbook_full_name.lower():
                    return wb
                if workbook_name and name == workbook_name.lower():
                    return wb
            except Exception:
                continue
        return None

    @staticmethod
    def parse_excel_link(link: str) -> Dict[str, Any]:
        """Parse a document-platform URL and extract lightweight workbook hints."""
        info = {"type": "unknown", "original_url": link.strip(), "file_name": None, "worksheet_hint": None, "error": None}
        try:
            clean = link.strip()
            if "sharepoint.com" in clean or "onedrive" in clean:
                info["type"] = "sharepoint"

            file_match = re.search(r"file=([^&]+)", clean, re.I)
            if file_match:
                info["file_name"] = urllib.parse.unquote(file_match.group(1))

            parsed = urllib.parse.urlparse(clean)
            query = urllib.parse.parse_qs(parsed.query)

            for key in ["sheet", "worksheet", "tab"]:
                if key in query and query[key]:
                    info["worksheet_hint"] = query[key][0]
                    break

            if not info["worksheet_hint"] and "activeCell" in query and query["activeCell"]:
                active_cell = query["activeCell"][0]
                if "!" in active_cell:
                    info["worksheet_hint"] = active_cell.split("!")[0]
        except Exception as e:
            info["error"] = str(e)
        return info

    def get_excel_connection_info(self) -> Dict[str, Any]:
        """Inspect current Excel runtime status for the UI status section."""
        self._coinit()
        try:
            info = {"excel_running": False, "workbooks": [], "active_workbook": None, "error": None}
            try:
                excel_app = win32.GetActiveObject("Excel.Application")
            except Exception as e:
                info["error"] = str(e)
                return info

            info["excel_running"] = True
            for i in range(1, excel_app.Workbooks.Count + 1):
                wb = excel_app.Workbooks(i)
                info["workbooks"].append({"name": str(wb.Name), "full_name": str(getattr(wb, "FullName", wb.Name))})

            if excel_app.Workbooks.Count > 0 and excel_app.ActiveWorkbook:
                info["active_workbook"] = {"name": str(excel_app.ActiveWorkbook.Name)}
            return info
        finally:
            self._couninit()

    def open_source(self, source: str, is_sharepoint: bool) -> bool:
        """Open source workbook either from document URL or local file."""
        return self._open_sharepoint(source) if is_sharepoint else self._open_local_file(source)

    def _open_sharepoint(self, link: str) -> bool:
        """Open a URL-backed workbook through desktop Excel and discover the workbook object."""
        self._coinit()
        try:
            link_info = self.parse_excel_link(link)
            expected_filename = link_info.get("file_name")
            worksheet_hint = link_info.get("worksheet_hint") or ""

            before_names = set()
            try:
                excel_before = win32.GetActiveObject("Excel.Application")
                for i in range(1, excel_before.Workbooks.Count + 1):
                    before_names.add(str(excel_before.Workbooks(i).Name))
            except Exception:
                pass

            self._log("🔗 Opening URL-backed workbook in Excel...")
            os.startfile(f"ms-excel:ofe|u|{link}")

            workbook = None
            for _ in range(20):
                time.sleep(2)
                try:
                    excel_app = win32.GetActiveObject("Excel.Application")
                    for i in range(1, excel_app.Workbooks.Count + 1):
                        wb = excel_app.Workbooks(i)
                        name = str(wb.Name)
                        if expected_filename and expected_filename.lower() in name.lower():
                            workbook = wb
                            break
                        if name not in before_names and not name.startswith("Book"):
                            workbook = wb
                            break
                    if workbook:
                        break
                except Exception:
                    pass

            if not workbook:
                self._log("❌ Failed to detect URL-backed workbook in Excel")
                return False

            active_sheet = ""
            try:
                active_sheet = str(workbook.ActiveSheet.Name)
            except Exception:
                pass

            worksheet_names = [str(workbook.Worksheets(i).Name) for i in range(1, workbook.Worksheets.Count + 1)]
            self.workbook_ref = WorkbookRef(
                workbook_name=str(workbook.Name),
                workbook_full_name=str(getattr(workbook, "FullName", workbook.Name)),
                workbook_path=str(getattr(workbook, "Path", "")),
                is_sharepoint=True,
                source=link,
                worksheet_names=worksheet_names,
                worksheet_hint=worksheet_hint,
                detected_active_sheet=active_sheet,
            )
            self._log(f"✅ Connected to workbook: {self.workbook_ref.workbook_name}")
            return True
        finally:
            self._couninit()

    def _open_local_file(self, file_path: str) -> bool:
        """Open or reconnect to a local workbook in Excel."""
        self._coinit()
        try:
            abs_path = os.path.abspath(file_path)
            self._log(f"📂 Opening local workbook: {abs_path}")
            excel_app = self._get_excel_app()
            excel_app.Visible = True

            wb = None
            for i in range(1, excel_app.Workbooks.Count + 1):
                candidate = excel_app.Workbooks(i)
                try:
                    candidate_full = str(getattr(candidate, "FullName", "")).lower()
                    candidate_name = str(getattr(candidate, "Name", "")).lower()
                    if candidate_full == abs_path.lower() or candidate_name == os.path.basename(abs_path).lower():
                        wb = candidate
                        break
                except Exception:
                    continue

            if wb is None:
                wb = excel_app.Workbooks.Open(abs_path)

            active_sheet = ""
            try:
                active_sheet = str(wb.ActiveSheet.Name)
            except Exception:
                pass

            worksheet_names = [str(wb.Worksheets(i).Name) for i in range(1, wb.Worksheets.Count + 1)]
            self.workbook_ref = WorkbookRef(
                workbook_name=str(wb.Name),
                workbook_full_name=str(getattr(wb, "FullName", wb.Name)),
                workbook_path=str(getattr(wb, "Path", "")),
                is_sharepoint=False,
                source=abs_path,
                worksheet_names=worksheet_names,
                detected_active_sheet=active_sheet,
            )
            self._log(f"✅ Local workbook connected: {self.workbook_ref.workbook_name}")
            return True
        finally:
            self._couninit()

    def suggest_default_worksheet(self) -> str:
        """Best-effort worksheet suggestion using URL hint, active sheet, or largest data range."""
        if not self.workbook_ref:
            return ""
        names = self.workbook_ref.worksheet_names

        if self.workbook_ref.worksheet_hint:
            for n in names:
                if n.lower() == self.workbook_ref.worksheet_hint.lower():
                    return n

        if self.workbook_ref.detected_active_sheet:
            for n in names:
                if n.lower() == self.workbook_ref.detected_active_sheet.lower():
                    return n

        worksheets = self.get_available_worksheets()
        if worksheets:
            return max(worksheets, key=lambda x: x["data_size"])["name"]

        return names[0] if names else ""

    def get_available_worksheets(self) -> List[Dict[str, Any]]:
        """Return worksheet metadata for picker population and suggestion logic."""
        if not self.workbook_ref:
            return []

        self._coinit()
        try:
            excel_app = self._get_excel_app()
            wb = self._find_workbook(excel_app, self.workbook_ref.workbook_name, self.workbook_ref.workbook_full_name)
            if wb is None:
                return []

            out = []
            for i in range(1, wb.Worksheets.Count + 1):
                ws = wb.Worksheets(i)
                try:
                    ur = ws.UsedRange
                    rows = int(ur.Rows.Count) if ur else 0
                    cols = int(ur.Columns.Count) if ur else 0
                    out.append({"name": str(ws.Name), "rows": rows, "cols": cols, "data_size": rows * cols, "has_data": rows >= 2})
                except Exception:
                    out.append({"name": str(ws.Name), "rows": 0, "cols": 0, "data_size": 0, "has_data": False})
            return out
        finally:
            self._couninit()

    def select_worksheet(self, worksheet_name: str):
        """Select active worksheet context and clear cached data."""
        self.current_worksheet_name = worksheet_name
        self.headers = []
        self.data_rows = []

    def _get_workbook_and_sheet(self):
        """Reacquire workbook and sheet COM objects on demand inside the calling thread."""
        if not self.workbook_ref:
            raise RuntimeError("No workbook connected")
        if not self.current_worksheet_name:
            raise RuntimeError("No worksheet selected")

        excel_app = self._get_excel_app()
        wb = self._find_workbook(excel_app, self.workbook_ref.workbook_name, self.workbook_ref.workbook_full_name)
        if wb is None:
            raise RuntimeError(f"Workbook not found: {self.workbook_ref.workbook_name}")
        ws = wb.Worksheets(self.current_worksheet_name)
        return excel_app, wb, ws

    def read_data_with_retry(self, max_retries: int = 3) -> bool:
        """Read worksheet UsedRange into headers + row cache with light retry protection."""
        last_error = None
        for attempt in range(1, max_retries + 1):
            self._coinit()
            try:
                self._log(f"📊 Reading worksheet data (attempt {attempt}/{max_retries})")
                _, _, ws = self._get_workbook_and_sheet()
                ur = ws.UsedRange
                if not ur:
                    raise RuntimeError("Worksheet has no used range")
                raw_data = ur.Value
                if not self._process_raw_data(raw_data):
                    raise RuntimeError("Failed to process worksheet data")
                self._log(f"✅ Loaded {len(self.data_rows)} data rows and {len(self.headers)} columns")
                return True
            except Exception as e:
                last_error = e
                self._log(f"⚠️ Worksheet read attempt failed: {e}")
                time.sleep(1.2)
            finally:
                self._couninit()
        raise RuntimeError(f"Failed to read data after retries: {last_error}")

    def _process_raw_data(self, raw_data) -> bool:
        """Normalize Excel UsedRange output into a consistent headers/rows structure."""
        if raw_data is None:
            return False

        if isinstance(raw_data, (list, tuple)) and len(raw_data) > 0:
            if isinstance(raw_data[0], (list, tuple)):
                rows = list(raw_data)
                self.headers = [str(v).strip() if v not in [None, ""] else f"Column_{i+1}" for i, v in enumerate(rows[0])]
                self.data_rows = [list(r) for r in rows[1:]]
            else:
                self.headers = [str(v).strip() if v not in [None, ""] else f"Column_{i+1}" for i, v in enumerate(raw_data)]
                self.data_rows = []
        else:
            self.headers = [str(raw_data)]
            self.data_rows = []

        return bool(self.headers)

    def ensure_work_item_ref_column(self, desired_column_name: str) -> str:
        """Ensure a reference column exists in row 1, creating it if missing."""
        self._coinit()
        try:
            _, wb, ws = self._get_workbook_and_sheet()
            ur = ws.UsedRange
            total_cols = int(ur.Columns.Count) if ur else 0

            for col in range(1, total_cols + 1):
                val = ws.Cells(1, col).Value
                if str(val or "").strip().lower() == desired_column_name.lower():
                    return str(val).strip()

            new_col = total_cols + 1 if total_cols > 0 else 1
            ws.Cells(1, new_col).Value = desired_column_name
            wb.Save()
            self._log(f"✅ Created work item reference column: {desired_column_name}")
            return desired_column_name
        finally:
            self._couninit()

    def write_work_item_ref(self, excel_row: int, work_item_ref_col: str, work_item_ref: str):
        """Write a created work-item ID back to Excel."""
        self._coinit()
        try:
            _, wb, ws = self._get_workbook_and_sheet()
            ur = ws.UsedRange
            total_cols = int(ur.Columns.Count) if ur else 0

            col_idx = None
            for col in range(1, total_cols + 1):
                header = str(ws.Cells(1, col).Value or "").strip()
                if header.lower() == work_item_ref_col.lower():
                    col_idx = col
                    break

            if col_idx is None:
                raise RuntimeError(f"Column not found: {work_item_ref_col}")

            ws.Cells(excel_row, col_idx).Value = work_item_ref
            wb.Save()
        finally:
            self._couninit()


# =============================================================================
# UI EVENT BUS
# =============================================================================

class UIEventBus:
    """Thread-safe event queue for marshalling worker updates back to Tk."""
    def __init__(self):
        self.q = queue.Queue()

    def emit(self, event_type: str, payload: Any = None):
        self.q.put((event_type, payload))


# =============================================================================
# FLOW UI
# =============================================================================

class WorkItemAutomationGUI:
    """Main desktop application controller and flow-oriented UI."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("1720x1000")
        self.root.minsize(1380, 840)

        self.event_bus = UIEventBus()
        self.config_obj = AppConfig()
        self.duplicate_protection = DuplicateProtection(SQLITE_DB)
        self.detector = IntelligentColumnDetector()
        self.analyzer = TaskAnalyzer(self.duplicate_protection, self.config_obj)
        self.work_item_service = WorkItemService()
        self.excel_service = ExcelService(ui_callback=self._emit_ui_event)

        self.column_mapping: Optional[ColumnMapping] = None
        self.task_proposals: List[TaskProposal] = []
        self.proposal_map: Dict[str, TaskProposal] = {}
        self.prepared_source: Optional[Dict[str, Any]] = None

        # Bridge Treeview item IDs back to proposal objects.
        self.row_item_to_proposal: Dict[str, TaskProposal] = {}

        # Mapping editor state.
        self.mapping_vars = {
            "status": tk.StringVar(value=""),
            "work_item_ref": tk.StringVar(value=""),
            "owner": tk.StringVar(value=""),
            "station_name": tk.StringVar(value=""),
            "description": tk.StringVar(value=""),
            "completed_date": tk.StringVar(value=""),
            "priority": tk.StringVar(value=""),
        }
        self.mapping_combo_boxes = {}

        # Selection editor state.
        self.edit_owner_var = tk.StringVar()
        self.edit_station_var = tk.StringVar()
        self.edit_priority_var = tk.StringVar()
        self.edit_status_var = tk.StringVar()
        self.edit_completed_var = tk.StringVar()

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.after(150, self._process_ui_events)
        self.root.after(500, self.refresh_excel_status_async)

    # -------------------------------------------------------------------------
    # Event bus
    # -------------------------------------------------------------------------

    def _emit_ui_event(self, event_type: str, payload: Any):
        self.event_bus.emit(event_type, payload)

    def _process_ui_events(self):
        """Drain background events safely on the main Tk thread."""
        try:
            while True:
                event_type, payload = self.event_bus.q.get_nowait()
                self._handle_ui_event(event_type, payload)
        except queue.Empty:
            pass
        self.root.after(150, self._process_ui_events)

    def _handle_ui_event(self, event_type: str, payload: Any):
        if event_type == "log":
            self._append_log(payload)
        elif event_type == "status":
            self.status_var.set(str(payload))
        elif event_type == "progress_max":
            self.progress.configure(maximum=int(payload), value=0)
        elif event_type == "progress_value":
            self.progress.configure(value=int(payload))
        elif event_type == "progress_label":
            self.progress_label.config(text=str(payload))
        elif event_type == "show_error":
            messagebox.showerror("Error", str(payload))
        elif event_type == "connected_success":
            self._on_connected_success()
        elif event_type == "data_read_success":
            self._on_data_read_success()
        elif event_type == "excel_status":
            self._update_excel_status_ui(payload)

    # -------------------------------------------------------------------------
    # UI construction
    # -------------------------------------------------------------------------

    def _build_ui(self):
        """Construct the 2-pane flow-based UI."""
        main = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        left = ttk.Frame(main, width=430)
        right = ttk.Frame(main)
        main.add(left, weight=0)
        main.add(right, weight=1)

        ttk.Label(left, text="🚀 Work Item Flow", font=("TkDefaultFont", 14, "bold")).pack(anchor=tk.W, padx=8, pady=(8, 4))

        self.step_var = tk.StringVar(value="Paste/Browse source → Open/Reopen workbook → Select worksheet → Load/Refresh data → Review/edit → Create work items")
        ttk.Label(left, textvariable=self.step_var, foreground="blue", wraplength=380).pack(anchor=tk.W, padx=8, pady=(0, 6))

        source_frame = ttk.LabelFrame(left, text="Source", padding=10)
        source_frame.pack(fill=tk.X, padx=8, pady=6)

        self.source_type_var = tk.StringVar(value="sharepoint")
        ttk.Radiobutton(source_frame, text="SharePoint / Excel URL", variable=self.source_type_var,
                        value="sharepoint", command=self.on_source_type_changed).pack(anchor=tk.W)
        ttk.Radiobutton(source_frame, text="Local Excel file", variable=self.source_type_var,
                        value="local", command=self.on_source_type_changed).pack(anchor=tk.W)

        self.sharepoint_frame = ttk.Frame(source_frame)
        self.sharepoint_frame.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(self.sharepoint_frame, text="Paste URL").pack(anchor=tk.W)
        self.excel_link_entry = tk.Text(self.sharepoint_frame, height=3, wrap=tk.WORD)
        self.excel_link_entry.pack(fill=tk.X, pady=(4, 6))
        self.excel_link_entry.bind("<Control-v>", self.on_link_paste)
        self.excel_link_entry.bind("<Control-V>", self.on_link_paste)
        self.excel_link_entry.bind("<Shift-Insert>", self.on_link_paste)
        self.excel_link_entry.bind("<Button-3>", self.show_paste_menu)

        sp_btns = ttk.Frame(self.sharepoint_frame)
        sp_btns.pack(fill=tk.X)
        ttk.Button(sp_btns, text="Paste URL", command=self.paste_from_clipboard).pack(side=tk.LEFT, padx=(0, 5))

        self.local_frame = ttk.Frame(source_frame)
        ttk.Label(self.local_frame, text="Choose local Excel file").pack(anchor=tk.W)
        local_btns = ttk.Frame(self.local_frame)
        local_btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(local_btns, text="Browse File", command=self.browse_local_file_only).pack(side=tk.LEFT)

        self.selected_local_file_var = tk.StringVar(value="")
        ttk.Label(self.local_frame, textvariable=self.selected_local_file_var, wraplength=360).pack(anchor=tk.W, pady=(6, 0))

        action_frame = ttk.LabelFrame(left, text="Actions", padding=10)
        action_frame.pack(fill=tk.X, padx=8, pady=6)

        self.prepared_summary_var = tk.StringVar(value="Source not prepared yet")
        ttk.Label(action_frame, textvariable=self.prepared_summary_var, wraplength=360, justify=tk.LEFT).pack(anchor=tk.W)

        btn_row1 = ttk.Frame(action_frame)
        btn_row1.pack(fill=tk.X, pady=(8, 0))
        self.open_workbook_button = ttk.Button(btn_row1, text="Open / Reopen Workbook in Excel", command=self.open_prepared_source, state="disabled")
        self.open_workbook_button.pack(side=tk.LEFT)

        worksheet_frame = ttk.LabelFrame(left, text="Worksheet", padding=10)
        worksheet_frame.pack(fill=tk.X, padx=8, pady=6)

        self.suggested_ws_var = tk.StringVar(value="Suggested worksheet: N/A")
        ttk.Label(worksheet_frame, textvariable=self.suggested_ws_var).pack(anchor=tk.W)

        self.worksheet_var = tk.StringVar()
        self.worksheet_combo = ttk.Combobox(worksheet_frame, textvariable=self.worksheet_var, state="readonly")
        self.worksheet_combo.pack(fill=tk.X, pady=(6, 6))
        self.worksheet_combo.bind("<<ComboboxSelected>>", self.on_worksheet_selected)

        ws_btns = ttk.Frame(worksheet_frame)
        ws_btns.pack(fill=tk.X)
        ttk.Button(ws_btns, text="Load / Refresh Worksheet Data", command=self.manual_read_data).pack(side=tk.LEFT)

        hitl_frame = ttk.LabelFrame(left, text="HITL Controls", padding=10)
        hitl_frame.pack(fill=tk.X, padx=8, pady=6)

        self.create_completed_only_var = tk.BooleanVar(value=self.config_obj.create_only_for_completed)
        self.auto_create_col_var = tk.BooleanVar(value=self.config_obj.auto_create_work_item_ref_column)
        self.confirm_open_var = tk.BooleanVar(value=self.config_obj.confirm_before_open_excel)
        self.confirm_create_var = tk.BooleanVar(value=self.config_obj.confirm_before_create_work_items)
        self.confirm_writeback_var = tk.BooleanVar(value=self.config_obj.confirm_before_writeback)

        ttk.Checkbutton(hitl_frame, text="Default recommend only completed tasks",
                        variable=self.create_completed_only_var, command=self.sync_config).pack(anchor=tk.W)
        ttk.Checkbutton(hitl_frame, text="Auto-create reference column if missing",
                        variable=self.auto_create_col_var, command=self.sync_config).pack(anchor=tk.W)
        ttk.Checkbutton(hitl_frame, text="Ask before opening/reopening workbook",
                        variable=self.confirm_open_var, command=self.sync_config).pack(anchor=tk.W)
        ttk.Checkbutton(hitl_frame, text="Ask before creating work items",
                        variable=self.confirm_create_var, command=self.sync_config).pack(anchor=tk.W)
        ttk.Checkbutton(hitl_frame, text="Ask before writeback to Excel",
                        variable=self.confirm_writeback_var, command=self.sync_config).pack(anchor=tk.W)

        excel_frame = ttk.LabelFrame(left, text="Excel Status", padding=10)
        excel_frame.pack(fill=tk.X, padx=8, pady=6)
        self.excel_status_label = ttk.Label(excel_frame, text="Checking Excel...")
        self.excel_status_label.pack(anchor=tk.W)
        self.workbook_status_label = ttk.Label(excel_frame, text="")
        self.workbook_status_label.pack(anchor=tk.W)

        log_frame = ttk.LabelFrame(left, text="Workflow Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=14, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        progress_frame = ttk.Frame(left)
        progress_frame.pack(fill=tk.X, padx=8, pady=(0, 8))
        self.progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress.pack(fill=tk.X)
        self.progress_label = ttk.Label(progress_frame, text="Ready")
        self.progress_label.pack(anchor=tk.W)

        # RIGHT WORKSPACE
        header = ttk.Frame(right)
        header.pack(fill=tk.X, padx=6, pady=6)
        ttk.Label(header, text="Workspace", font=("TkDefaultFont", 14, "bold")).pack(side=tk.LEFT)

        stats = ttk.Frame(right)
        stats.pack(fill=tk.X, padx=6, pady=(0, 6))
        self.summary_var = tk.StringVar(value="No data loaded")
        ttk.Label(stats, textvariable=self.summary_var, foreground="green").pack(anchor=tk.W)

        upper = ttk.PanedWindow(right, orient=tk.VERTICAL)
        upper.pack(fill=tk.BOTH, expand=True)

        grid_frame = ttk.LabelFrame(upper, text="Task Grid", padding=6)
        upper.add(grid_frame, weight=4)

        self.data_tree = ttk.Treeview(grid_frame, show="headings", selectmode="extended")
        data_v = ttk.Scrollbar(grid_frame, orient=tk.VERTICAL, command=self.data_tree.yview)
        data_h = ttk.Scrollbar(grid_frame, orient=tk.HORIZONTAL, command=self.data_tree.xview)
        self.data_tree.configure(yscrollcommand=data_v.set, xscrollcommand=data_h.set)

        self.data_tree.grid(row=0, column=0, sticky="nsew")
        data_v.grid(row=0, column=1, sticky="ns")
        data_h.grid(row=1, column=0, sticky="ew")

        grid_frame.grid_rowconfigure(0, weight=1)
        grid_frame.grid_columnconfigure(0, weight=1)
        self.data_tree.bind("<<TreeviewSelect>>", self.on_tree_selection_changed)

        lower = ttk.PanedWindow(upper, orient=tk.HORIZONTAL)
        upper.add(lower, weight=2)

        mapping_frame = ttk.LabelFrame(lower, text="Column Mapping", padding=8)
        lower.add(mapping_frame, weight=1)

        field_order = [
            ("status", "Status"),
            ("work_item_ref", "Work Item Ref"),
            ("owner", "Owner"),
            ("station_name", "Station"),
            ("description", "Description"),
            ("completed_date", "Completed Date"),
            ("priority", "Priority"),
        ]

        for row_idx, (field_key, label_text) in enumerate(field_order):
            ttk.Label(mapping_frame, text=label_text).grid(row=row_idx, column=0, sticky="w", padx=4, pady=3)
            combo = ttk.Combobox(mapping_frame, textvariable=self.mapping_vars[field_key], state="readonly", width=34)
            combo.grid(row=row_idx, column=1, sticky="ew", padx=4, pady=3)
            self.mapping_combo_boxes[field_key] = combo

        mapping_frame.grid_columnconfigure(1, weight=1)

        map_btns = ttk.Frame(mapping_frame)
        map_btns.grid(row=len(field_order), column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(map_btns, text="Apply Mapping", command=self.apply_manual_mapping).pack(side=tk.LEFT, padx=3)
        ttk.Button(map_btns, text="Re-detect", command=self.redetect_mapping).pack(side=tk.LEFT, padx=3)
        ttk.Button(map_btns, text="Refresh Proposals", command=self.generate_proposals).pack(side=tk.LEFT, padx=3)

        editor_frame = ttk.LabelFrame(lower, text="Selection Editor & Actions", padding=8)
        lower.add(editor_frame, weight=1)

        ttk.Label(editor_frame, text="Owner").grid(row=0, column=0, sticky="w", padx=4, pady=3)
        ttk.Entry(editor_frame, textvariable=self.edit_owner_var).grid(row=0, column=1, sticky="ew", padx=4, pady=3)

        ttk.Label(editor_frame, text="Station").grid(row=1, column=0, sticky="w", padx=4, pady=3)
        ttk.Entry(editor_frame, textvariable=self.edit_station_var).grid(row=1, column=1, sticky="ew", padx=4, pady=3)

        ttk.Label(editor_frame, text="Priority").grid(row=2, column=0, sticky="w", padx=4, pady=3)
        ttk.Entry(editor_frame, textvariable=self.edit_priority_var).grid(row=2, column=1, sticky="ew", padx=4, pady=3)

        ttk.Label(editor_frame, text="Status").grid(row=3, column=0, sticky="w", padx=4, pady=3)
        ttk.Entry(editor_frame, textvariable=self.edit_status_var).grid(row=3, column=1, sticky="ew", padx=4, pady=3)

        ttk.Label(editor_frame, text="Completed Date").grid(row=4, column=0, sticky="w", padx=4, pady=3)
        ttk.Entry(editor_frame, textvariable=self.edit_completed_var).grid(row=4, column=1, sticky="ew", padx=4, pady=3)

        ttk.Label(editor_frame, text="Description").grid(row=5, column=0, sticky="nw", padx=4, pady=3)
        self.edit_description_text = tk.Text(editor_frame, height=6, wrap=tk.WORD)
        self.edit_description_text.grid(row=5, column=1, sticky="nsew", padx=4, pady=3)

        editor_frame.grid_columnconfigure(1, weight=1)
        editor_frame.grid_rowconfigure(5, weight=1)

        sel_btns = ttk.Frame(editor_frame)
        sel_btns.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(8, 4))
        ttk.Button(sel_btns, text="Apply Edit to Selected", command=self.apply_edit_to_selected).pack(side=tk.LEFT, padx=3)
        ttk.Button(sel_btns, text="Select All", command=self.select_all_rows).pack(side=tk.LEFT, padx=3)
        ttk.Button(sel_btns, text="Select Recommended", command=self.select_recommended_rows).pack(side=tk.LEFT, padx=3)
        ttk.Button(sel_btns, text="Clear Selection", command=self.clear_selection).pack(side=tk.LEFT, padx=3)

        action_btns = ttk.Frame(editor_frame)
        action_btns.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(4, 4))
        ttk.Button(action_btns, text="Approve Selected", command=self.approve_selected).pack(side=tk.LEFT, padx=3)
        ttk.Button(action_btns, text="Skip Selected", command=self.skip_selected).pack(side=tk.LEFT, padx=3)
        ttk.Button(action_btns, text="Approve High Confidence", command=self.approve_high_confidence).pack(side=tk.LEFT, padx=3)

        run_btns = ttk.Frame(editor_frame)
        run_btns.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        ttk.Button(run_btns, text="Dry Run", command=lambda: self.execute_work_items(True)).pack(side=tk.RIGHT, padx=3)
        ttk.Button(run_btns, text="Create Work Items", command=lambda: self.execute_work_items(False)).pack(side=tk.RIGHT, padx=3)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN).pack(side=tk.BOTTOM, fill=tk.X)

        self.on_source_type_changed()

    # -------------------------------------------------------------------------
    # Helpers
    # -------------------------------------------------------------------------

    def sync_config(self):
        """Sync checkbox state into runtime config."""
        self.config_obj.create_only_for_completed = self.create_completed_only_var.get()
        self.config_obj.auto_create_work_item_ref_column = self.auto_create_col_var.get()
        self.config_obj.confirm_before_open_excel = self.confirm_open_var.get()
        self.config_obj.confirm_before_create_work_items = self.confirm_create_var.get()
        self.config_obj.confirm_before_writeback = self.confirm_writeback_var.get()

    def _append_log(self, message: str):
        """Append a timestamped workflow message."""
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] {message}\n")
        self.log_text.see(tk.END)

    def on_source_type_changed(self):
        """Show source controls relevant to the current source mode."""
        if self.source_type_var.get() == "sharepoint":
            self.local_frame.pack_forget()
            self.sharepoint_frame.pack(fill=tk.X, pady=(8, 0))
        else:
            self.sharepoint_frame.pack_forget()
            self.local_frame.pack(fill=tk.X, pady=(8, 0))

    # -------------------------------------------------------------------------
    # Source
    # -------------------------------------------------------------------------

    def paste_clipboard_to_link_box(self, auto_prepare=True):
        """Paste URL from clipboard and optionally auto-prepare it."""
        try:
            text = self.root.clipboard_get().strip()
            try:
                self.excel_link_entry.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            self.excel_link_entry.insert(tk.INSERT, text)
            self._append_log("📋 URL pasted from clipboard")
            if auto_prepare:
                self.root.after(100, self.prepare_source)
        except Exception as e:
            messagebox.showerror("Paste Error", f"Could not paste from clipboard:\n{e}")

    def paste_from_clipboard(self):
        """Explicit paste action for URL mode."""
        self.paste_clipboard_to_link_box(auto_prepare=True)

    def on_link_paste(self, event=None):
        """Keyboard paste hook for Ctrl+V / Shift+Insert."""
        self.paste_clipboard_to_link_box(auto_prepare=True)
        return "break"

    def show_paste_menu(self, event):
        """Minimal right-click context menu for URL entry."""
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Paste", command=self.paste_from_clipboard)
        menu.tk_popup(event.x_root, event.y_root)

    def browse_local_file_only(self):
        """Browse for a local Excel file and auto-prepare it."""
        file_path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")]
        )
        if file_path:
            self.selected_local_file_var.set(file_path)
            self._append_log(f"📂 Selected local file: {os.path.basename(file_path)}")
            self.prepare_source()

    def prepare_source(self):
        """
        Prepare source metadata automatically.

        This step is intentionally low-friction; the actual workbook open/reopen
        remains an explicit human action.
        """
        if self.source_type_var.get() == "sharepoint":
            source = self.excel_link_entry.get("1.0", tk.END).strip()
            if not source:
                return
            parsed = self.excel_service.parse_excel_link(source)
            self.prepared_source = {"source": source, "is_sharepoint": True, "parsed": parsed}
            summary = (
                f"Prepared SharePoint source\n"
                f"File: {parsed.get('file_name') or 'Unknown'}\n"
                f"Worksheet hint: {parsed.get('worksheet_hint') or 'N/A'}"
            )
        else:
            source = self.selected_local_file_var.get().strip()
            if not source:
                return
            self.prepared_source = {"source": source, "is_sharepoint": False, "parsed": {"file_name": os.path.basename(source)}}
            summary = f"Prepared local file\nFile: {os.path.basename(source)}"

        self.prepared_summary_var.set(summary)
        self.open_workbook_button.config(state="normal")
        self.status_var.set("Source prepared. Click 'Open / Reopen Workbook in Excel'.")
        self._append_log("✅ Source prepared automatically")

    # -------------------------------------------------------------------------
    # Workbook / worksheet
    # -------------------------------------------------------------------------

    def open_prepared_source(self):
        """Open or reconnect workbook in desktop Excel with explicit confirmation."""
        if not self.prepared_source:
            messagebox.showerror("Error", "Please provide a source first")
            return

        source = self.prepared_source["source"]
        is_sharepoint = self.prepared_source["is_sharepoint"]

        if self.config_obj.confirm_before_open_excel:
            if not messagebox.askyesno("Open / Reopen Workbook", f"This will open or reconnect the workbook in Excel.\n\n{source}\n\nContinue?"):
                return

        self.status_var.set("Opening/reopening workbook...")
        threading.Thread(target=self._open_source_worker, args=(source, is_sharepoint), daemon=True).start()

    def _open_source_worker(self, source: str, is_sharepoint: bool):
        """Background workbook open/reconnect worker."""
        try:
            ok = self.excel_service.open_source(source, is_sharepoint)
            if ok:
                self.event_bus.emit("connected_success", None)
            else:
                self.event_bus.emit("show_error", "Failed to open workbook")
        except Exception as e:
            self.event_bus.emit("show_error", str(e))

    def _on_connected_success(self):
        """Post-open sync: refresh status, discover worksheets, suggest default sheet."""
        self._append_log("✅ Workbook opened and connected")
        self.status_var.set("Workbook connected. Select worksheet or load data.")
        self.refresh_excel_status_async()
        self.refresh_worksheets()

        suggested = self.excel_service.suggest_default_worksheet()
        if suggested:
            self.suggested_ws_var.set(f"Suggested worksheet: {suggested}")
            self.worksheet_var.set(suggested)
            self.excel_service.select_worksheet(suggested)

    def refresh_excel_status_async(self):
        """Refresh passive Excel status display shown in the dedicated status section."""
        def worker():
            try:
                info = self.excel_service.get_excel_connection_info()
                self.event_bus.emit("excel_status", info)
            except Exception as e:
                self.event_bus.emit("excel_status", {"excel_running": False, "error": str(e)})
        threading.Thread(target=worker, daemon=True).start()

    def _update_excel_status_ui(self, info: Dict[str, Any]):
        """Render current Excel process/workbook status."""
        if info.get("excel_running"):
            count = len(info.get("workbooks", []))
            self.excel_status_label.config(text=f"✅ Excel running ({count} workbooks)")
            active = info.get("active_workbook")
            self.workbook_status_label.config(text=f"Active workbook: {active['name']}" if active else "")
        else:
            self.excel_status_label.config(text="❌ Excel not running")
            self.workbook_status_label.config(text=info.get("error", ""))

    def refresh_worksheets(self):
        """Populate worksheet picker and select the best default suggestion if possible."""
        try:
            worksheets = self.excel_service.get_available_worksheets()
            names = [w["name"] for w in worksheets]
            self.worksheet_combo["values"] = names

            suggested = self.excel_service.suggest_default_worksheet()
            if suggested and suggested in names:
                self.worksheet_var.set(suggested)
                self.excel_service.select_worksheet(suggested)
                self.suggested_ws_var.set(f"Suggested worksheet: {suggested}")
        except Exception as e:
            self._append_log(f"❌ Refresh worksheet error: {e}")

    def on_worksheet_selected(self, event=None):
        """Track current worksheet selection."""
        selected = self.worksheet_var.get()
        if selected:
            self.excel_service.select_worksheet(selected)
            self._append_log(f"📋 Worksheet selected: {selected}")

    def manual_read_data(self):
        """Load or refresh data from the selected worksheet."""
        if not self.excel_service.current_worksheet_name:
            messagebox.showerror("Error", "Please select a worksheet")
            return
        self.status_var.set("Loading / refreshing worksheet data...")
        threading.Thread(target=self._read_data_worker, daemon=True).start()

    def _read_data_worker(self):
        """Background worksheet read worker."""
        try:
            self.excel_service.read_data_with_retry()
            self.event_bus.emit("data_read_success", None)
        except Exception as e:
            self.event_bus.emit("show_error", f"Data load failed:\n{e}")

    def _on_data_read_success(self):
        """
        Continue the normal flow automatically after read:
        - detect mapping
        - generate proposals
        - refresh workspace
        """
        self.column_mapping = self.detector.detect_columns(self.excel_service.headers, self.excel_service.data_rows)
        self.populate_mapping_dropdowns()
        self.generate_proposals(auto_log=False)
        self.refresh_grid()
        self.update_summary()
        self._append_log("✅ Worksheet data refreshed, mapping detected, proposals generated automatically")
        self.status_var.set("Data ready. Review, edit, select, and create.")

    # -------------------------------------------------------------------------
    # Mapping
    # -------------------------------------------------------------------------

    def populate_mapping_dropdowns(self):
        """Populate mapping selectors from current worksheet headers."""
        columns = [""] + self.excel_service.headers
        for combo in self.mapping_combo_boxes.values():
            combo["values"] = columns

        if self.column_mapping:
            self.mapping_vars["status"].set(self.column_mapping.status or "")
            self.mapping_vars["work_item_ref"].set(self.column_mapping.work_item_ref or "")
            self.mapping_vars["owner"].set(self.column_mapping.owner or "")
            self.mapping_vars["station_name"].set(self.column_mapping.station_name or "")
            self.mapping_vars["description"].set(self.column_mapping.description or "")
            self.mapping_vars["completed_date"].set(self.column_mapping.completed_date or "")
            self.mapping_vars["priority"].set(self.column_mapping.priority or "")

    def apply_manual_mapping(self):
        """Commit current mapping selections and rebuild proposals."""
        prev_conf = self.column_mapping.confidence if self.column_mapping else 0.0
        self.column_mapping = ColumnMapping(
            status=self.mapping_vars["status"].get() or None,
            work_item_ref=self.mapping_vars["work_item_ref"].get() or None,
            owner=self.mapping_vars["owner"].get() or None,
            station_name=self.mapping_vars["station_name"].get() or None,
            description=self.mapping_vars["description"].get() or None,
            completed_date=self.mapping_vars["completed_date"].get() or None,
            priority=self.mapping_vars["priority"].get() or None,
            confidence=prev_conf
        )
        self.generate_proposals(auto_log=False)
        self.refresh_grid()
        self.update_summary()
        self._append_log("✅ Manual mapping applied and proposals refreshed")

    def redetect_mapping(self):
        """Re-run heuristic mapping and refresh proposals immediately."""
        if not self.excel_service.headers:
            messagebox.showerror("Error", "Load worksheet data first")
            return
        self.column_mapping = self.detector.detect_columns(self.excel_service.headers, self.excel_service.data_rows)
        self.populate_mapping_dropdowns()
        self.generate_proposals(auto_log=False)
        self.refresh_grid()
        self.update_summary()
        self._append_log("🔄 Re-detected mapping suggestions and refreshed proposals")

    def generate_proposals(self, auto_log=True):
        """Rebuild all proposals from current worksheet data + current mapping."""
        if not self.column_mapping:
            return

        self.task_proposals = []
        self.proposal_map = {}

        for i, row_data in enumerate(self.excel_service.data_rows):
            proposal = self.analyzer.analyze_row(row_data, self.column_mapping, self.excel_service.headers, i)
            self.task_proposals.append(proposal)
            self.proposal_map[proposal.proposal_id] = proposal

        if auto_log:
            self._append_log(f"✅ Generated {len(self.task_proposals)} proposals")

    # -------------------------------------------------------------------------
    # Grid / selection / editing
    # -------------------------------------------------------------------------

    def refresh_grid(self):
        """Render proposals into the main workspace grid."""
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)
        self.row_item_to_proposal.clear()

        cols = (
            "Excel Row", "Action", "Owner", "Station", "Station Code", "Priority",
            "Status", "Completed", "Description", "Existing Ref", "Confidence", "Reason"
        )
        self.data_tree["columns"] = cols

        widths = {
            "Excel Row": 80,
            "Action": 90,
            "Owner": 120,
            "Station": 170,
            "Station Code": 90,
            "Priority": 80,
            "Status": 100,
            "Completed": 110,
            "Description": 300,
            "Existing Ref": 140,
            "Confidence": 90,
            "Reason": 220,
        }

        for col in cols:
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=widths.get(col, 120), stretch=True)

        for proposal in self.task_proposals:
            desc = proposal.description[:80] + "..." if len(proposal.description) > 80 else proposal.description
            reason = proposal.decision_reason[:80] + "..." if len(proposal.decision_reason) > 80 else proposal.decision_reason
            item = self.data_tree.insert("", tk.END, values=(
                proposal.excel_row,
                proposal.recommended_action.title(),
                proposal.owner,
                proposal.station_name,
                proposal.station_code,
                proposal.priority,
                proposal.status,
                proposal.completed_date,
                desc,
                proposal.existing_work_item_id or "",
                f"{proposal.confidence:.1%}",
                reason,
            ), tags=(proposal.recommended_action,))
            self.row_item_to_proposal[item] = proposal

        self.data_tree.tag_configure("create", background="#e8f5e8")
        self.data_tree.tag_configure("skip", background="#f5f5f5")

    def on_tree_selection_changed(self, event=None):
        """Load single selected row into the side editor; clear editor on multi-select."""
        selected = self._get_selected_proposals()
        if len(selected) == 1:
            p = selected[0]
            self.edit_owner_var.set(p.owner)
            self.edit_station_var.set(p.station_name)
            self.edit_priority_var.set(p.priority)
            self.edit_status_var.set(p.status)
            self.edit_completed_var.set(p.completed_date)
            self.edit_description_text.delete("1.0", tk.END)
            self.edit_description_text.insert("1.0", p.description)
        elif len(selected) > 1:
            self.edit_owner_var.set("")
            self.edit_station_var.set("")
            self.edit_priority_var.set("")
            self.edit_status_var.set("")
            self.edit_completed_var.set("")
            self.edit_description_text.delete("1.0", tk.END)

    def apply_edit_to_selected(self):
        """Apply side-editor values to all currently selected rows."""
        selected = self._get_selected_proposals()
        if not selected:
            messagebox.showinfo("Info", "Select one or more rows first")
            return

        owner = self.edit_owner_var.get().strip()
        station = self.edit_station_var.get().strip()
        priority = self.edit_priority_var.get().strip()
        status = self.edit_status_var.get().strip()
        completed = self.edit_completed_var.get().strip()
        description = self.edit_description_text.get("1.0", tk.END).strip()

        for p in selected:
            if owner:
                p.owner = owner
                p.source_data["owner"] = owner
            if station:
                p.station_name = station
                p.station_code = self.analyzer.extract_station_code(station)
                p.source_data["station_name"] = station
            if priority:
                p.priority = priority
                p.source_data["priority"] = priority
            if status:
                p.status = status
                p.source_data["status"] = status
            if completed:
                p.completed_date = completed
                p.source_data["completed_date"] = completed
            if description:
                p.description = description
                p.source_data["description"] = description

            p.task_hash = self.duplicate_protection.compute_task_hash(p.source_data)

        for p in selected:
            has_existing = bool(p.existing_work_item_id)
            is_dup, dup_id = self.duplicate_protection.is_duplicate(p.task_hash)
            recommendation = "create"
            reasons = []

            if has_existing:
                recommendation = "skip"
                reasons.append("existing work item ref found")
            elif is_dup and dup_id != p.existing_work_item_id:
                recommendation = "skip"
                reasons.append("duplicate task hash found")
            else:
                if not p.station_name:
                    reasons.append("missing station")
                if not p.description or len(p.description) < self.config_obj.min_description_length:
                    reasons.append("short/missing description")
                if self.config_obj.create_only_for_completed and not any(s in p.status.lower() for s in COMPLETED_STATUSES):
                    recommendation = "skip"
                    reasons.append("not completed while completed-only mode is enabled")
                else:
                    if not p.status:
                        reasons.append("missing status allowed")
                    elif "open" in p.status.lower():
                        reasons.append("open status allowed")

            p.recommended_action = recommendation
            p.decision_reason = "; ".join(reasons) if reasons else "eligible for creation"

        self.refresh_grid()
        self.update_summary()
        self._append_log(f"✏️ Applied edits to {len(selected)} selected row(s)")

    def update_summary(self):
        """Update compact summary text above the workspace."""
        total = len(self.task_proposals)
        create_count = sum(1 for p in self.task_proposals if p.recommended_action == "create")
        skip_count = sum(1 for p in self.task_proposals if p.recommended_action == "skip")
        self.summary_var.set(
            f"Rows loaded: {len(self.excel_service.data_rows)} | Proposals: {total} | Create: {create_count} | Skip: {skip_count}"
        )

    # -------------------------------------------------------------------------
    # Bulk selection / actions
    # -------------------------------------------------------------------------

    def select_all_rows(self):
        """Select all visible rows in the grid."""
        items = self.data_tree.get_children()
        self.data_tree.selection_set(items)

    def select_recommended_rows(self):
        """Select only rows currently recommended for creation."""
        items = []
        for item, proposal in self.row_item_to_proposal.items():
            if proposal.recommended_action == "create":
                items.append(item)
        self.data_tree.selection_set(items)

    def clear_selection(self):
        """Clear the current grid selection."""
        self.data_tree.selection_remove(self.data_tree.selection())

    def _get_selected_proposals(self) -> List[TaskProposal]:
        """Resolve Treeview selection into proposal objects."""
        selected = []
        for item in self.data_tree.selection():
            p = self.row_item_to_proposal.get(item)
            if p:
                selected.append(p)
        return selected

    def approve_selected(self):
        """Force selected rows into create state."""
        selected = self._get_selected_proposals()
        for p in selected:
            p.recommended_action = "create"
            p.decision_reason = "manually approved"
        self.refresh_grid()
        self.update_summary()
        self._append_log(f"✅ Approved {len(selected)} selected row(s)")

    def skip_selected(self):
        """Force selected rows into skip state."""
        selected = self._get_selected_proposals()
        for p in selected:
            p.recommended_action = "skip"
            p.decision_reason = "manually skipped"
        self.refresh_grid()
        self.update_summary()
        self._append_log(f"⏭️ Marked {len(selected)} selected row(s) to skip")

    def approve_high_confidence(self):
        """Bulk-approve rows above the configured confidence threshold."""
        count = 0
        for p in self.task_proposals:
            if p.confidence > HIGH_CONFIDENCE_THRESHOLD and not p.has_existing_work_item:
                p.recommended_action = "create"
                p.decision_reason = "auto-approved high confidence"
                count += 1
        self.refresh_grid()
        self.update_summary()
        self._append_log(f"🎯 Approved {count} high-confidence row(s)")

    # -------------------------------------------------------------------------
    # Execution
    # -------------------------------------------------------------------------

    def execute_work_items(self, dry_run: bool):
        """Execute approved proposals as dry run or mock live creation."""
        approved = [p for p in self.task_proposals if p.recommended_action == "create"]
        if not approved:
            messagebox.showinfo("Info", "No approved rows to process")
            return

        if not dry_run and self.config_obj.confirm_before_create_work_items:
            if not messagebox.askyesno("Create Work Items", f"{len(approved)} work items are approved.\nContinue?"):
                return

        if not dry_run and self.config_obj.confirm_before_writeback:
            if not messagebox.askyesno(
                "Write Back to Excel",
                f"Created work-item IDs will be written back to worksheet '{self.excel_service.current_worksheet_name}'.\nContinue?"
            ):
                return

        threading.Thread(target=self._execute_worker, args=(approved, dry_run), daemon=True).start()

    def _execute_worker(self, proposals: List[TaskProposal], dry_run: bool):
        """Background execution worker for dry run or mock live create/writeback."""
        self.event_bus.emit("progress_max", len(proposals))
        self.event_bus.emit("log", f"🚀 Starting {'DRY RUN' if dry_run else 'LIVE RUN'} for {len(proposals)} items")

        work_item_ref_col = None
        if not dry_run:
            if self.column_mapping and self.column_mapping.work_item_ref:
                work_item_ref_col = self.column_mapping.work_item_ref
            elif self.config_obj.auto_create_work_item_ref_column:
                work_item_ref_col = self.excel_service.ensure_work_item_ref_column(self.config_obj.work_item_ref_column_name)
                if self.column_mapping:
                    self.column_mapping.work_item_ref = work_item_ref_col

        for i, proposal in enumerate(proposals, start=1):
            try:
                self.event_bus.emit("progress_label", f"Processing {i}/{len(proposals)}")

                if dry_run:
                    payload = self.work_item_service.build_work_item_payload(proposal)
                    self.event_bus.emit("log", f"🔍 DRY RUN row {proposal.excel_row}: {payload['title']}")
                else:
                    work_item_id, payload = self.work_item_service.create_work_item(proposal)
                    self.duplicate_protection.record_processed_task(proposal.task_hash, work_item_id, proposal.source_data)

                    if work_item_ref_col:
                        self.excel_service.write_work_item_ref(proposal.excel_row, work_item_ref_col, work_item_id)

                    proposal.has_existing_work_item = True
                    proposal.existing_work_item_id = work_item_id
                    proposal.recommended_action = "skip"
                    proposal.decision_reason = "work item created"

                    self.event_bus.emit("log", f"✅ Created {work_item_id} for row {proposal.excel_row}")
                    self.event_bus.emit("log", f"   Title: {payload['title']}")
                    self.event_bus.emit("log", f"   Owner: {payload['owner'] or 'N/A'} | Priority: {payload['priority'] or 'N/A'} | Station: {payload['station_name'] or 'N/A'}")

                self.event_bus.emit("progress_value", i)

            except Exception as e:
                self.event_bus.emit("log", f"❌ Failed for row {proposal.excel_row}: {e}")

        self.event_bus.emit("progress_label", "Complete")
        self.event_bus.emit("status", "Execution complete")
        self.root.after(0, self.refresh_grid)
        self.root.after(0, self.update_summary)

    # -------------------------------------------------------------------------
    # Shutdown
    # -------------------------------------------------------------------------

    def on_closing(self):
        """Close the application."""
        self.root.destroy()

    def run(self):
        """Start the Tk main loop."""
        self.root.mainloop()


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Application entry point."""
    print("🚀 Starting Spreadsheet Work Item Automation - Flow UI")
    print("📂 Supports local Excel files and SharePoint URLs")
    print("🧭 Flow-based workspace")
    print("🗺️ Manual mapping + auto proposals")
    print("🧪 Dry run and mock live execution")
    print("🪟 Windows + Excel + pywin32 required")

    try:
        app = WorkItemAutomationGUI()
        app.run()
    except Exception as e:
        logger.exception("Fatal application error")
        print(f"Fatal error: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
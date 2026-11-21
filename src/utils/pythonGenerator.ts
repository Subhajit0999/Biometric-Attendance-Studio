import type { AttendanceConfig, PythonFeatureFlags } from '../types/attendance';

export const DEFAULT_PYTHON_FLAGS: PythonFeatureFlags = {
  bakeCurrentDefaults: true,
  bakeCurrentHolidays: true,
  includeGaussianTiming: true,
  includeLateAndHalfDay: true,
  includeComplianceGuard: true,
  includeFlexibleWeekends: true,
  includeStyledExcelAndSummary: true,
  includeCsvAndHtmlExport: true,
  includeTabbedScrollableGui: true,
  includeCliMode: true,
  includeJsonProfiles: true,
};

function pyStr(val: string): string {
  return JSON.stringify(val ?? '');
}

export function generateRequirementsTxt(flags: PythonFeatureFlags): string {
  const deps = [
    'pandas>=2.1.0',
    'openpyxl>=3.1.2',
  ];
  if (flags.includeTabbedScrollableGui) {
    deps.push('tkcalendar>=1.6.1');
  }
  return deps.join('\n') + '\n';
}

export function generateDynamicPythonCode(
  config: AttendanceConfig,
  flags: PythonFeatureFlags
): string {
  const defaultCenter = flags.bakeCurrentDefaults ? config.centerName : '';
  const defaultDistrict = flags.bakeCurrentDefaults ? config.district : '';
  const defaultLocation = flags.bakeCurrentDefaults ? config.location : '';
  const defaultBatch = flags.bakeCurrentDefaults ? config.batchNo : '';
  const defaultQual = flags.bakeCurrentDefaults ? config.qualificationPack : '';
  const defaultScheme = flags.bakeCurrentDefaults ? config.schemeName : '';
  const defaultTrainer = flags.bakeCurrentDefaults ? config.trainerName : '';
  const defaultManager = flags.bakeCurrentDefaults ? config.centerManager : '';
  const defaultInMin = flags.bakeCurrentDefaults ? config.inMin : '10:00';
  const defaultInMax = flags.bakeCurrentDefaults ? config.inMax : '10:30';
  const defaultOutMin = flags.bakeCurrentDefaults ? config.outMin : '16:00';
  const defaultOutMax = flags.bakeCurrentDefaults ? config.outMax : '16:30';
  const defaultTimeFmt = flags.bakeCurrentDefaults ? config.timeFormat : '12h_short';
  const defaultWeekend = flags.bakeCurrentDefaults ? config.weekendPolicy : 'sunday';
  const defaultAbsentPolicy = flags.bakeCurrentDefaults ? config.absentPolicy : 'range';
  const defaultCompliance = flags.bakeCurrentDefaults ? config.complianceThresholdPct : 75;
  const defaultSeed = flags.bakeCurrentDefaults ? config.seed : 2026;
  const defaultTheme = flags.bakeCurrentDefaults ? config.excelTheme : 'classic_navy';

  const holidaysLiteral =
    flags.bakeCurrentHolidays && config.holidays.length > 0
      ? config.holidays
          .map((h) => `        HolidayEntry(date_iso=${pyStr(h.date)}, label=${pyStr(h.label)}),`)
          .join('\n')
      : '        # Add default holidays here, e.g. HolidayEntry("2026-08-15", "Independence Day")';

  const enabledFeaturesSummary = [
    flags.includeGaussianTiming ? 'Gaussian Bell-Curve Timing + Shift Duration Guard' : 'Uniform Timing',
    flags.includeLateAndHalfDay ? 'Late Arrival & Half-Day Simulation' : null,
    flags.includeComplianceGuard ? 'Target % & Audit Compliance Floor Guard' : null,
    flags.includeFlexibleWeekends ? 'Flexible Weekend Policies (Sun / Sat+Sun / 2nd&4th Sat)' : 'Sunday-only Weekends',
    flags.includeStyledExcelAndSummary ? 'Styled OpenPyXL + Executive Summary Sheet + Print/Freeze' : 'Clean Bordered OpenPyXL',
    flags.includeCsvAndHtmlExport ? 'CSV + Printable HTML Report Exporters' : null,
    flags.includeJsonProfiles ? 'JSON Configuration Profile Save/Load' : null,
    flags.includeTabbedScrollableGui ? 'Modern Scrollable Tabbed Tkinter GUI + Live Preview' : 'Headless Engine',
    flags.includeCliMode ? 'Full CLI Mode (argparse)' : null,
  ]
    .filter(Boolean)
    .map((f) => `#   - ${f}`)
    .join('\n');

  const sections: string[] = [];

  // =========================================================================
  // SECTION 1: HEADER, IMPORTS & CONFIGURATION DATACLASSES
  // =========================================================================
  sections.push(`#!/usr/bin/env python3
"""
Biometric Attendance Maker Pro (Dynamic Edition)
================================================
Upgraded, modular, production-grade biometric attendance generator.

Enabled Modules:
${enabledFeaturesSummary}

Quick Start:
    pip install -r requirements.txt
    python attendance_maker_pro.py          # Launch Desktop GUI
    python attendance_maker_pro.py --help   # CLI automation options
"""

from __future__ import annotations

import argparse
import calendar
import csv
import glob
import html as html_lib
import json
import logging
import math
import os
import random
import re
import threading
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, time as dt_time, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
${
  flags.includeTabbedScrollableGui
    ? `
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkcalendar import DateEntry
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False
`
    : 'GUI_AVAILABLE = False\n'
}
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("AttendanceMakerPro")

# =============================================================================
# SECTION 1 — CONFIGURATION MODELS & CONSTANTS
# =============================================================================

CANDIDATE_HEADER_KEYWORDS: Tuple[str, ...] = (
    "candidate",
    "name",
    "student",
    "trainee",
    "participant",
    "candidate name",
    "student name",
    "full name",
)

INDIAN_HOLIDAY_PRESETS: List[Tuple[str, str]] = [
    ("2026-01-26", "Republic Day"),
    ("2026-03-04", "Holi"),
    ("2026-03-21", "Id-ul-Fitr"),
    ("2026-04-03", "Good Friday"),
    ("2026-04-14", "Ambedkar Jayanti"),
    ("2026-05-01", "Labour Day"),
    ("2026-08-15", "Independence Day"),
    ("2026-10-02", "Gandhi Jayanti"),
    ("2026-10-20", "Dussehra"),
    ("2026-11-08", "Diwali"),
    ("2026-11-24", "Guru Nanak Jayanti"),
    ("2026-12-25", "Christmas Day"),
]

SHIFT_PRESETS: Dict[str, Tuple[str, str, str, str, str]] = {
    "Standard (10:00-16:30)": ("10:00", "10:30", "16:00", "16:30", "13:15"),
    "Morning (09:00-13:30)": ("09:00", "09:20", "13:15", "13:40", "11:15"),
    "Afternoon (13:30-18:00)": ("13:20", "13:45", "17:50", "18:15", "15:45"),
    "Full Day (09:30-17:30)": ("09:15", "09:45", "17:20", "17:50", "13:30"),
}

EXCEL_THEMES: Dict[str, Dict[str, str]] = {
    "classic_navy": {
        "banner_bg": "1E3A5F", "banner_fg": "FFFFFF",
        "sub_bg": "EFF6FF", "sub_fg": "0F172A",
        "hdr_bg": "1E293B", "hdr_fg": "FFFFFF",
        "wk_bg": "E2E8F0", "wk_fg": "334155",
        "border": "CBD5E1", "zebra": "F8FAFC",
    },
    "slate_minimal": {
        "banner_bg": "0F172A", "banner_fg": "FFFFFF",
        "sub_bg": "F1F5F9", "sub_fg": "0F172A",
        "hdr_bg": "334155", "hdr_fg": "FFFFFF",
        "wk_bg": "E2E8F0", "wk_fg": "334155",
        "border": "CBD5E1", "zebra": "F8FAFC",
    },
    "emerald_gov": {
        "banner_bg": "065F46", "banner_fg": "FFFFFF",
        "sub_bg": "ECFDF5", "sub_fg": "064E3B",
        "hdr_bg": "047857", "hdr_fg": "FFFFFF",
        "wk_bg": "D1FAE5", "wk_fg": "065F46",
        "border": "A7F3D0", "zebra": "F0FDF4",
    },
    "monochrome": {
        "banner_bg": "FFFFFF", "banner_fg": "000000",
        "sub_bg": "FFFFFF", "sub_fg": "000000",
        "hdr_bg": "F1F5F9", "hdr_fg": "000000",
        "wk_bg": "F8FAFC", "wk_fg": "334155",
        "border": "000000", "zebra": "FFFFFF",
    },
}


@dataclass
class HolidayEntry:
    date_iso: str  # YYYY-MM-DD
    label: str = "Holiday"

    def to_date(self) -> date:
        return datetime.strptime(self.date_iso, "%Y-%m-%d").date()


@dataclass
class CandidateRecord:
    name: str
    roll_no: str = ""
    join_date_iso: Optional[str] = None
    drop_date_iso: Optional[str] = None
    forced_absents: List[str] = field(default_factory=list)


@dataclass
class AttendanceConfig:
    # Center & Batch Metadata
    center_name: str = ${pyStr(defaultCenter)}
    district: str = ${pyStr(defaultDistrict)}
    location: str = ${pyStr(defaultLocation)}
    batch_no: str = ${pyStr(defaultBatch)}
    qualification_pack: str = ${pyStr(defaultQual)}
    scheme_name: str = ${pyStr(defaultScheme)}
    trainer_name: str = ${pyStr(defaultTrainer)}
    center_manager: str = ${pyStr(defaultManager)}
    report_title: str = "Monthly Attendance Report"

    # Mode & Date Range
    mode: str = ${pyStr(config.mode)}  # single | months_stacked | months_separate
    start_date: str = ${pyStr(config.startDate)}
    end_date: str = ${pyStr(config.endDate)}
    selected_months: List[int] = field(default_factory=lambda: ${JSON.stringify(config.selectedMonths)})

    # Weekend Policy
    weekend_policy: str = ${pyStr(defaultWeekend)}  # sunday | sat_sun | second_fourth_sat | custom
    custom_weekend_days: List[int] = field(default_factory=lambda: [6])  # Python weekday: 0=Mon..6=Sun

    # Biometric Timing Rules (24h HH:MM internally)
    auto_fill_times: bool = ${config.autoFillTimes ? 'True' : 'False'}
    in_min: str = ${pyStr(defaultInMin)}
    in_max: str = ${pyStr(defaultInMax)}
    out_min: str = ${pyStr(defaultOutMin)}
    out_max: str = ${pyStr(defaultOutMax)}
    time_format: str = ${pyStr(defaultTimeFmt)}  # 12h_short | 12h_ampm | 24h
    time_distribution: str = ${pyStr(flags.includeGaussianTiming ? config.timeDistribution : 'uniform')}
    min_shift_hours: float = ${config.minShiftHours}

    # Late & Half-Day Simulation
    late_enabled: bool = ${flags.includeLateAndHalfDay && config.lateSimulationEnabled ? 'True' : 'False'}
    late_chance_pct: float = ${config.lateChancePercent}
    late_max_extra_mins: int = ${config.lateMaxExtraMinutes}
    half_day_enabled: bool = ${flags.includeLateAndHalfDay && config.halfDaySimulationEnabled ? 'True' : 'False'}
    half_day_chance_pct: float = ${config.halfDayChancePercent}
    half_day_out_time: str = ${pyStr(config.halfDayOutTime)}

    # Absent & Audit Compliance Rules
    absent_policy: str = ${pyStr(defaultAbsentPolicy)}  # none | fixed | range | target_percent
    fixed_absents_per_month: int = ${config.fixedAbsentsPerMonth}
    min_absents_per_month: int = ${config.minAbsentsPerMonth}
    max_absents_per_month: int = ${config.maxAbsentsPerMonth}
    target_min_attendance_pct: float = ${config.targetMinAttendancePct}
    target_max_attendance_pct: float = ${config.targetMaxAttendancePct}
    compliance_threshold_pct: float = ${defaultCompliance}
    avoid_consecutive_absents: bool = ${config.avoidConsecutiveAbsents ? 'True' : 'False'}
    absent_marker: str = ${pyStr(config.absentMarker)}

    # Reproducibility & Holidays
    seed: Optional[int] = ${defaultSeed}
    holidays: List[HolidayEntry] = field(default_factory=lambda: [
${holidaysLiteral}
    ])

    # Excel Styling & Output Options
    excel_theme: str = ${pyStr(defaultTheme)}
    color_code_cells: bool = ${flags.includeStyledExcelAndSummary && config.colorCodeCells ? 'True' : 'False'}
    include_summary_sheet: bool = ${flags.includeStyledExcelAndSummary && config.includeSummarySheet ? 'True' : 'False'}
    freeze_panes: bool = ${config.freezePanes ? 'True' : 'False'}
    landscape_print: bool = ${config.landscapePrintSetup ? 'True' : 'False'}
    include_signature_block: bool = ${config.includeSignatureBlock ? 'True' : 'False'}
${
  flags.includeJsonProfiles
    ? `
    def save_profile(self, path: str | Path) -> None:
        """Persist configuration profile to JSON."""
        data = asdict(self)
        Path(path).write_text(json.dumps(data, indent=2), encoding="utf-8")
        logger.info("Saved configuration profile to %s", path)

    @classmethod
    def load_profile(cls, path: str | Path) -> "AttendanceConfig":
        """Load configuration profile from JSON with backward-compatible defaults."""
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
        holidays = [HolidayEntry(**h) for h in raw.pop("holidays", [])]
        valid_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in raw.items() if k in valid_fields}
        return cls(**filtered, holidays=holidays)
`
    : ''
}`);

  // =========================================================================
  // SECTION 2: SMART FILE & ROSTER LOADER
  // =========================================================================
  sections.push(`
# =============================================================================
# SECTION 2 — SMART ROSTER & FILE LOADER
# =============================================================================

def get_script_directory() -> Path:
    """Return the folder where this Python script is running."""
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def find_roster_files(folder: str | Path = ".") -> List[str]:
    """Return sorted list of Excel/CSV candidate files in folder."""
    exts = ("*.xlsx", "*.xls", "*.xlsm", "*.csv")
    found: List[str] = []
    for pattern in exts:
        found.extend(glob.glob(os.path.join(str(folder), pattern)))
    # Exclude previously generated attendance outputs and Excel temp lock files (~$)
    filtered = [
        f for f in sorted(set(found))
        if "_Attendance_" not in os.path.basename(f)
        and not os.path.basename(f).startswith("~$")
    ]
    return filtered


def auto_fetch_roster_from_folder(
    folder: Optional[str | Path] = None,
    column_override: Optional[str] = None,
    deduplicate: bool = True,
) -> Tuple[List[CandidateRecord], Optional[str], Optional[str]]:
    """
    Auto-fetch the best Excel/CSV candidate file from the given folder
    (defaults to the directory where this Python script is running).
    Returns: (candidates, best_file_path, chosen_column)
    """
    search_dirs: List[Path] = []
    if folder is not None:
        search_dirs.append(Path(folder))
    else:
        script_dir = get_script_directory()
        search_dirs.append(script_dir)
        cwd_dir = Path.cwd().resolve()
        if cwd_dir != script_dir:
            search_dirs.append(cwd_dir)

    best_cands: List[CandidateRecord] = []
    best_file: Optional[str] = None
    best_col: Optional[str] = None

    for directory in search_dirs:
        for file_path in find_roster_files(directory):
            try:
                cands, col, _ = read_roster_file(
                    file_path,
                    column_override=column_override,
                    deduplicate=deduplicate,
                )
                if len(cands) > len(best_cands):
                    best_cands, best_file, best_col = cands, file_path, col
            except Exception as exc:
                logger.warning("Skipping unreadable roster file %s: %s", file_path, exc)
        if best_cands:
            break

    return best_cands, best_file, best_col


def sniff_candidate_column(df: pd.DataFrame) -> Optional[str]:
    """Auto-detect the candidate name column using header keywords & heuristics."""
    if df is None or df.shape[1] == 0:
        return None
    for col in df.columns:
        norm = str(col).strip().lower()
        if any(kw in norm for kw in CANDIDATE_HEADER_KEYWORDS):
            return str(col)
    first = str(df.columns[0]).strip().lower()
    if any(x in first for x in ("sl", "s.no", "s/no", "sl no", "slno", "sr")):
        return str(df.columns[1] if df.shape[1] >= 2 else df.columns[0])
    return str(df.columns[0])


def sniff_roll_column(df: pd.DataFrame, name_col: str) -> Optional[str]:
    """Detect optional Roll No / Registration ID column."""
    for col in df.columns:
        if str(col) == name_col:
            continue
        norm = str(col).strip().lower()
        if any(k in norm for k in ("roll", "reg", "id", "code", "enrol")):
            return str(col)
    return None


def read_roster_file(
    path: str | Path,
    column_override: Optional[str] = None,
    deduplicate: bool = True,
) -> Tuple[List[CandidateRecord], Optional[str], List[str]]:
    """
    Read candidates from Excel or CSV safely with zero bare excepts.
    Returns: (candidates, chosen_column, all_columns)
    """
    path_str = str(path)
    if path_str.lower().endswith(".csv"):
        df = pd.read_csv(path_str, dtype=str)
    else:
        try:
            df = pd.read_excel(path_str, dtype=str, engine="openpyxl")
        except Exception as exc:
            logger.warning("openpyxl read fallback for %s: %s", path_str, exc)
            df = pd.read_excel(path_str, dtype=str)

    df = df.dropna(how="all")
    all_cols = [str(c) for c in df.columns.tolist()]
    if df.empty or not all_cols:
        return [], None, all_cols

    chosen = column_override if (column_override and column_override in all_cols) else sniff_candidate_column(df)
    if not chosen:
        return [], None, all_cols

    roll_col = sniff_roll_column(df, chosen)
    candidates: List[CandidateRecord] = []
    seen: Set[str] = set()

    for _, row in df.iterrows():
        raw_name = str(row.get(chosen, "")).strip()
        if not raw_name or raw_name.lower() in ("nan", "none"):
            continue
        if raw_name.lower() == chosen.strip().lower():
            continue
        key = raw_name.lower()
        if deduplicate and key in seen:
            continue
        seen.add(key)

        roll_val = ""
        if roll_col:
            rv = str(row.get(roll_col, "")).strip()
            if rv.lower() not in ("nan", "none"):
                roll_val = rv

        candidates.append(CandidateRecord(name=raw_name, roll_no=roll_val))

    return candidates, chosen, all_cols


@dataclass
class PreviousAttendanceData:
    """Parsed timestamps and candidates from a previously generated attendance Excel/CSV."""
    file_name: str
    candidates: List[CandidateRecord]
    records: Dict[str, Dict[date, Tuple[str, str]]]
    min_date: Optional[date]
    max_date: Optional[date]
    dates_count: int
    detected_meta: Dict[str, str] = field(default_factory=dict)


def _parse_attendance_header_date(raw_val) -> Optional[date]:
    """Parse date headers such as '01-04-2026', '01-04-2026 (Wed)', '2026-04-01', or Excel date cells."""
    if raw_val is None:
        return None
    if isinstance(raw_val, datetime):
        return raw_val.date()
    if isinstance(raw_val, date):
        return raw_val
    text = str(raw_val).strip()
    if not text:
        return None
    # Strip trailing weekday annotation like '01-04-2026 (Wed)'
    text = re.sub(r"\\s*\\([^)]*\\)\\s*$", "", text).strip()
    m_dmy = re.match(r"^(\\d{1,2})[-/](\\d{1,2})[-/](\\d{4})$", text)
    if m_dmy:
        try:
            return date(int(m_dmy.group(3)), int(m_dmy.group(2)), int(m_dmy.group(1)))
        except ValueError:
            return None
    m_ymd = re.match(r"^(\\d{4})[-/](\\d{1,2})[-/](\\d{1,2})$", text)
    if m_ymd:
        try:
            return date(int(m_ymd.group(1)), int(m_ymd.group(2)), int(m_ymd.group(3)))
        except ValueError:
            return None
    return None


def _cell_to_punch_str(raw_val) -> str:
    """Convert an openpyxl/pandas cell value to a clean punch text string."""
    if raw_val is None:
        return ""
    if isinstance(raw_val, datetime):
        return raw_val.strftime("%H:%M")
    if isinstance(raw_val, dt_time):
        return raw_val.strftime("%H:%M")
    s = str(raw_val).strip()
    if s.lower() in ("nan", "none"):
        return ""
    return s


def parse_previous_attendance_excel(path: str | Path) -> PreviousAttendanceData:
    """
    Parse a previously generated attendance Excel (.xlsx/.xls) or CSV workbook
    so earlier IN/OUT timestamps are preserved verbatim when extending to new dates/months.
    Supports Single Range, Stacked Multi-Month, and Separate Monthly Sheet formats.
    """
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"Previous attendance file not found: {file_path}")

    sheets_rows: List[List[List[object]]] = []
    suffix = file_path.suffix.lower()

    if suffix == ".csv":
        with file_path.open("r", encoding="utf-8-sig", newline="") as fh:
            sheets_rows.append([row for row in csv.reader(fh)])
    elif suffix in (".xlsx", ".xlsm"):
        wb = load_workbook(file_path, data_only=True)
        try:
            for name in wb.sheetnames:
                if name.strip().lower() == "executive_summary":
                    continue
                ws = wb[name]
                rows_list: List[List[object]] = []
                for row in ws.iter_rows(values_only=True):
                    rows_list.append(list(row))
                sheets_rows.append(rows_list)
        finally:
            wb.close()
    else:
        xls_dict = pd.read_excel(file_path, sheet_name=None, header=None, dtype=object)
        for name, df in xls_dict.items():
            if str(name).strip().lower() == "executive_summary":
                continue
            rows_list = []
            for _, row in df.iterrows():
                rows_list.append([None if pd.isna(v) else v for v in row.tolist()])
            sheets_rows.append(rows_list)

    candidates: List[CandidateRecord] = []
    seen_candidates: Set[str] = set()
    records_mut: Dict[str, Dict[date, List[str]]] = {}
    all_dates: Set[date] = set()
    detected_meta: Dict[str, str] = {}

    for sheet_rows in sheets_rows:
        active_date_cols: List[Tuple[int, date]] = []
        current_name = ""

        for row in sheet_rows:
            if not row or all(v is None or str(v).strip() == "" for v in row):
                continue

            # Extract optional center metadata from banner lines
            first_text = _cell_to_punch_str(row[0]) if len(row) > 0 else ""
            if not active_date_cols and first_text:
                if "(Dist:" in first_text or "(Batch:" in first_text:
                    center_part = re.split(r"\\s*\\((?:Dist|Batch):", first_text, maxsplit=1)[0].strip()
                    if center_part and "center_name" not in detected_meta:
                        detected_meta["center_name"] = center_part
                    m_dist = re.search(r"\\(Dist:\\s*([^)]+)\\)", first_text)
                    if m_dist and "district" not in detected_meta:
                        detected_meta["district"] = m_dist.group(1).strip()
                    m_batch = re.search(r"\\(Batch:\\s*([^)]+)\\)", first_text)
                    if m_batch and "batch_no" not in detected_meta:
                        detected_meta["batch_no"] = m_batch.group(1).strip()
                elif "Location:" in first_text:
                    m_loc = re.search(r"Location:\\s*(.+)$", first_text)
                    if m_loc and "location" not in detected_meta:
                        detected_meta["location"] = m_loc.group(1).strip()
                elif first_text.startswith("Qualification Pack:"):
                    qp = first_text.replace("Qualification Pack:", "", 1).strip()
                    if qp and "qualification_pack" not in detected_meta:
                        detected_meta["qualification_pack"] = qp

            # Check if this row is a date header row (Sl No | Candidate Name | Type | DD-MM-YYYY ...)
            parsed_cols: List[Tuple[int, date]] = []
            for col_idx in range(3, len(row)):
                dt = _parse_attendance_header_date(row[col_idx])
                if dt is not None:
                    parsed_cols.append((col_idx, dt))

            col2_low = _cell_to_punch_str(row[2]).lower() if len(row) > 2 else ""
            col1_low = _cell_to_punch_str(row[1]).lower() if len(row) > 1 else ""
            if len(parsed_cols) >= 2 and ("type" in col2_low or "candidate" in col1_low):
                active_date_cols = parsed_cols
                current_name = ""
                for _, dt in parsed_cols:
                    all_dates.add(dt)
                continue

            if not active_date_cols:
                continue

            # Detect signature footer or new block banner
            if first_text.upper().startswith("BIOMETRIC ATTENDANCE SHEET") or (
                len(row) > 1 and _cell_to_punch_str(row[1]).startswith("Trainer:")
            ):
                active_date_cols = []
                current_name = ""
                continue

            row_type = _cell_to_punch_str(row[2]).lower() if len(row) > 2 else ""
            if row_type not in ("in", "out", "punch in", "punch out"):
                continue

            raw_name = _cell_to_punch_str(row[1]) if len(row) > 1 else ""
            if raw_name:
                current_name = raw_name
            if not current_name:
                continue

            key = current_name.strip().lower()
            if key not in seen_candidates:
                seen_candidates.add(key)
                candidates.append(CandidateRecord(name=current_name.strip()))

            cand_map = records_mut.setdefault(key, {})
            is_in = row_type in ("in", "punch in")
            for col_idx, dt in active_date_cols:
                val = _cell_to_punch_str(row[col_idx]) if col_idx < len(row) else ""
                pair = cand_map.get(dt)
                if pair is None:
                    pair = ["", ""]
                    cand_map[dt] = pair
                if is_in:
                    pair[0] = val
                else:
                    pair[1] = val

    if not candidates or not all_dates:
        raise ValueError(
            "Could not find valid attendance dates and candidate In/Out rows in the selected file."
        )

    records_final: Dict[str, Dict[date, Tuple[str, str]]] = {
        k: {dt: (pair[0], pair[1]) for dt, pair in dmap.items()}
        for k, dmap in records_mut.items()
    }
    sorted_dates = sorted(all_dates)
    return PreviousAttendanceData(
        file_name=file_path.name,
        candidates=candidates,
        records=records_final,
        min_date=sorted_dates[0],
        max_date=sorted_dates[-1],
        dates_count=len(sorted_dates),
        detected_meta=detected_meta,
    )
`);

  // =========================================================================
  // SECTION 3: BIOMETRIC TIMING & ATTENDANCE ENGINE
  // =========================================================================
  sections.push(`
# =============================================================================
# SECTION 3 — DYNAMIC BIOMETRIC ATTENDANCE ENGINE
# =============================================================================

@dataclass
class DayPunch:
    day: date
    status: str  # present | late | half_day | absent | weekend | holiday | inactive | blank
    in_str: str
    out_str: str
    duration_mins: int = 0
    holiday_label: str = ""
    preserved: bool = False


@dataclass
class CandidateBlockResult:
    sl_no: int
    candidate: CandidateRecord
    punches: Dict[date, DayPunch]
    working_days: int
    present_days: float
    absent_days: int
    late_days: int
    half_days: int
    total_hours: float
    attendance_pct: float
    compliant: bool


@dataclass
class AttendanceBlock:
    sheet_name: str
    month_label: str
    span_start: date
    span_end: date
    dates: List[date]
    holiday_map: Dict[date, str]
    weekend_dates: Set[date]
    rows: List[CandidateBlockResult]


def parse_hhmm_to_minutes(hhmm: str) -> int:
    """Parse HH:MM (24h) or HH:MM AM/PM into minutes since midnight."""
    s = hhmm.strip().upper()
    for fmt in ("%H:%M", "%I:%M %p", "%I:%M%p"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.hour * 60 + dt.minute
        except ValueError:
            continue
    raise ValueError(f"Invalid time format '{hhmm}'. Expected HH:MM (e.g. 10:00 or 16:30).")


def format_minutes(total_mins: int, fmt: str) -> str:
    """Format minutes since midnight as 12h_short, 12h_ampm, or 24h."""
    clamped = total_mins % 1440
    h24, m = divmod(clamped, 60)
    if fmt == "24h":
        return f"{h24:02d}:{m:02d}"
    ampm = "PM" if h24 >= 12 else "AM"
    h12 = 12 if h24 % 12 == 0 else h24 % 12
    if fmt == "12h_ampm":
        return f"{h12:02d}:{m:02d} {ampm}"
    return f"{h12:02d}:{m:02d}"


def sample_minutes(min_m: int, max_m: int, distribution: str, rng: random.Random) -> int:
    """Pick biometric punch minute using Gaussian clustering or uniform sampling."""
    lo, hi = (min_m, max_m) if min_m <= max_m else (max_m, min_m)
    if hi <= lo:
        return lo
${
  flags.includeGaussianTiming
    ? `    if distribution == "gaussian":
        mean = lo + (hi - lo) * 0.45
        std = max(1.0, (hi - lo) / 4.2)
        val = int(round(rng.gauss(mean, std)))
        return max(lo, min(hi, val))`
    : ''
}
    return rng.randint(lo, hi)


def is_weekend_day(d: date, cfg: AttendanceConfig) -> bool:
    """Evaluate weekend policy (Python weekday: 0=Mon .. 5=Sat, 6=Sun)."""
    wd = d.weekday()
${
  flags.includeFlexibleWeekends
    ? `    if cfg.weekend_policy == "sat_sun":
        return wd in (5, 6)
    if cfg.weekend_policy == "second_fourth_sat":
        if wd == 6:
            return True
        if wd == 5:
            nth_sat = math.ceil(d.day / 7)
            return nth_sat in (2, 4)
        return False
    if cfg.weekend_policy == "custom":
        return wd in set(cfg.custom_weekend_days)`
    : ''
}
    return wd == 6


def _pick_absent_dates(
    working_dates: Sequence[date],
    count: int,
    avoid_consecutive: bool,
    rng: random.Random,
) -> Set[date]:
    target = max(0, min(count, len(working_dates)))
    if target == 0:
        return set()
    indices = list(range(len(working_dates)))
    rng.shuffle(indices)
    if not avoid_consecutive:
        return {working_dates[i] for i in indices[:target]}

    picked: List[int] = []
    for idx in indices:
        if len(picked) >= target:
            break
        if all(abs(idx - p) != 1 for p in picked):
            picked.append(idx)
    if len(picked) < target:
        for idx in indices:
            if len(picked) >= target:
                break
            if idx not in picked:
                picked.append(idx)
    return {working_dates[i] for i in picked}


def _is_marker_value(val: str, absent_marker: str) -> bool:
    s = val.strip().upper()
    return s in ("", "A", "ABSENT", "AB", "OFF", "WO", "H", "HOLIDAY", "-", "--", absent_marker.strip().upper())


def _parse_punch_to_minutes(val: str, is_out: bool = False) -> Optional[int]:
    s = val.strip().upper()
    if not s:
        return None
    m = re.match(r"^(\\d{1,2}):(\\d{2})(?:\\s*(AM|PM))?$", s)
    if not m:
        return None
    hh = int(m.group(1))
    mm = int(m.group(2))
    ampm = m.group(3)
    if ampm == "PM" and hh < 12:
        hh += 12
    elif ampm == "AM" and hh == 12:
        hh = 0
    elif ampm is None and is_out and 1 <= hh <= 7:
        hh += 12
    return hh * 60 + mm


class AttendanceEngine:
    """Core deterministic biometric attendance generator with continuation support."""

    def __init__(
        self,
        cfg: AttendanceConfig,
        previous_records: Optional[Dict[str, Dict[date, Tuple[str, str]]]] = None,
    ):
        self.cfg = cfg
        self.previous_records = previous_records or {}
        self.rng = random.Random(cfg.seed) if cfg.seed is not None else random.Random()

    def build_blocks(self, candidates: Sequence[CandidateRecord]) -> List[AttendanceBlock]:
        if not candidates:
            raise ValueError("Candidate list is empty.")

        start_d = datetime.strptime(self.cfg.start_date, "%Y-%m-%d").date()
        end_d = datetime.strptime(self.cfg.end_date, "%Y-%m-%d").date()
        if end_d < start_d:
            raise ValueError("End date must be on or after start date.")

        if self.cfg.mode == "single":
            m_label = (
                start_d.strftime("%B %Y")
                if (start_d.year == end_d.year and start_d.month == end_d.month)
                else f"{start_d.strftime('%b %Y')} - {end_d.strftime('%b %Y')}"
            )
            return [self._build_span("Attendance_ALL", m_label, start_d, end_d, candidates)]

        if not self.cfg.selected_months:
            raise ValueError("Select at least one month for By-Months mode.")

        ym_list: List[Tuple[int, int]] = []
        cur = date(start_d.year, start_d.month, 1)
        last = date(end_d.year, end_d.month, 1)
        while cur <= last:
            ym_list.append((cur.year, cur.month))
            cur = date(cur.year + 1, 1, 1) if cur.month == 12 else date(cur.year, cur.month + 1, 1)

        blocks: List[AttendanceBlock] = []
        for yr, mo in ym_list:
            if mo not in self.cfg.selected_months:
                continue
            first_m = date(yr, mo, 1)
            last_m = date(yr, mo, calendar.monthrange(yr, mo)[1])
            s_d = max(first_m, start_d)
            e_d = min(last_m, end_d)
            if e_d < s_d:
                continue
            m_name = calendar.month_name[mo]
            sheet_name = f"{m_name[:3]}_{yr}" if self.cfg.mode == "months_separate" else "Attendance_ALL"
            blocks.append(self._build_span(sheet_name, f"{m_name} {yr}", s_d, e_d, candidates))

        if not blocks:
            raise ValueError("Selected months do not overlap the start/end date span.")
        return blocks

    def _build_span(
        self,
        sheet_name: str,
        month_label: str,
        span_start: date,
        span_end: date,
        candidates: Sequence[CandidateRecord],
    ) -> AttendanceBlock:
        cfg = self.cfg
        holiday_map: Dict[date, str] = {}
        for h in cfg.holidays:
            try:
                holiday_map[h.to_date()] = h.label or "Holiday"
            except ValueError:
                continue

        dates: List[date] = []
        cur = span_start
        while cur <= span_end:
            dates.append(cur)
            cur += timedelta(days=1)

        weekend_dates = {d for d in dates if is_weekend_day(d, cfg)}
        working_by_ym: Dict[Tuple[int, int], List[date]] = {}
        for d in dates:
            if d in weekend_dates or d in holiday_map:
                continue
            working_by_ym.setdefault((d.year, d.month), []).append(d)

        in_min_m = parse_hhmm_to_minutes(cfg.in_min)
        in_max_m = parse_hhmm_to_minutes(cfg.in_max)
        out_min_m = parse_hhmm_to_minutes(cfg.out_min)
        out_max_m = parse_hhmm_to_minutes(cfg.out_max)
        half_out_m = parse_hhmm_to_minutes(cfg.half_day_out_time)
        min_shift_mins = int(round(max(0.0, cfg.min_shift_hours) * 60))

        rows: List[CandidateBlockResult] = []
        for idx, cand in enumerate(candidates, start=1):
            join_d = datetime.strptime(cand.join_date_iso, "%Y-%m-%d").date() if cand.join_date_iso else None
            drop_d = datetime.strptime(cand.drop_date_iso, "%Y-%m-%d").date() if cand.drop_date_iso else None
            forced_set = {
                datetime.strptime(ds, "%Y-%m-%d").date()
                for ds in cand.forced_absents
                if ds
            }
            cand_prev = self.previous_records.get(cand.name.strip().lower(), {})

            auto_absents: Set[date] = set()
            for _, wdates in working_by_ym.items():
                eligible = [
                    d for d in wdates
                    if (join_d is None or d >= join_d) and (drop_d is None or d <= drop_d)
                ]
                forced_in_mo = sum(1 for d in eligible if d in forced_set and d not in cand_prev)
                preserved_abs_in_mo = sum(
                    1
                    for d in eligible
                    if d in cand_prev
                    and _is_marker_value(cand_prev[d][0], cfg.absent_marker)
                    and _is_marker_value(cand_prev[d][1], cfg.absent_marker)
                )
                avail = [d for d in eligible if d not in forced_set and d not in cand_prev]

                desired = 0
                if cfg.absent_policy == "fixed":
                    desired = cfg.fixed_absents_per_month
                elif cfg.absent_policy == "range":
                    lo = min(cfg.min_absents_per_month, cfg.max_absents_per_month)
                    hi = max(cfg.min_absents_per_month, cfg.max_absents_per_month)
                    desired = self.rng.randint(lo, hi)
${
  flags.includeComplianceGuard
    ? `                elif cfg.absent_policy == "target_percent":
                    lo_p = min(cfg.target_min_attendance_pct, cfg.target_max_attendance_pct)
                    hi_p = max(cfg.target_min_attendance_pct, cfg.target_max_attendance_pct)
                    chosen_pct = self.rng.uniform(lo_p, hi_p)
                    desired = int(round(len(eligible) * (1.0 - chosen_pct / 100.0)))
                    max_allowed = int(math.floor(len(eligible) * (1.0 - cfg.compliance_threshold_pct / 100.0)))
                    desired = min(desired, max(0, max_allowed))`
    : ''
}
                need = max(0, desired - forced_in_mo - preserved_abs_in_mo)
                auto_absents |= _pick_absent_dates(avail, need, cfg.avoid_consecutive_absents, self.rng)

            punches: Dict[date, DayPunch] = {}
            working_days = 0
            present_days = 0.0
            absent_days = 0
            late_days = 0
            half_days = 0
            total_mins = 0

            for d in dates:
                if d in cand_prev:
                    prev_in, prev_out = cand_prev[d]
                    is_prev_abs = _is_marker_value(prev_in, cfg.absent_marker) and _is_marker_value(prev_out, cfg.absent_marker)
                    if is_prev_abs:
                        if d in weekend_dates:
                            punches[d] = DayPunch(d, "weekend", prev_in, prev_out, preserved=True)
                        elif d in holiday_map:
                            punches[d] = DayPunch(d, "holiday", prev_in, prev_out, holiday_label=holiday_map[d], preserved=True)
                        elif (join_d and d < join_d) or (drop_d and d > drop_d):
                            punches[d] = DayPunch(d, "inactive", prev_in, prev_out, preserved=True)
                        else:
                            working_days += 1
                            absent_days += 1
                            punches[d] = DayPunch(d, "absent", prev_in, prev_out, preserved=True)
                    else:
                        working_days += 1
                        present_days += 1.0
                        in_p = _parse_punch_to_minutes(prev_in, is_out=False)
                        out_p = _parse_punch_to_minutes(prev_out, is_out=True)
                        dur = (out_p - in_p) if (in_p is not None and out_p is not None and out_p > in_p) else 0
                        total_mins += dur
                        punches[d] = DayPunch(d, "present", prev_in, prev_out, duration_mins=dur, preserved=True)
                    continue

                if d in weekend_dates:
                    punches[d] = DayPunch(d, "weekend", cfg.absent_marker, cfg.absent_marker)
                elif d in holiday_map:
                    punches[d] = DayPunch(d, "holiday", cfg.absent_marker, cfg.absent_marker, holiday_label=holiday_map[d])
                elif (join_d and d < join_d) or (drop_d and d > drop_d):
                    punches[d] = DayPunch(d, "inactive", "-", "-")
                elif d in forced_set or d in auto_absents:
                    working_days += 1
                    absent_days += 1
                    punches[d] = DayPunch(d, "absent", cfg.absent_marker, cfg.absent_marker)
                elif not cfg.auto_fill_times:
                    working_days += 1
                    punches[d] = DayPunch(d, "blank", "", "")
                else:
                    working_days += 1
                    is_late = ${flags.includeLateAndHalfDay ? 'cfg.late_enabled and (self.rng.random() * 100.0 < cfg.late_chance_pct)' : 'False'}
                    is_half = ${flags.includeLateAndHalfDay ? '(not is_late) and cfg.half_day_enabled and (self.rng.random() * 100.0 < cfg.half_day_chance_pct)' : 'False'}

                    if is_late:
                        extra = self.rng.randint(3, max(3, cfg.late_max_extra_mins))
                        in_m = min(1430, max(in_min_m, in_max_m) + extra)
                        out_m = sample_minutes(out_min_m, out_max_m, cfg.time_distribution, self.rng)
                        status = "late"
                        present_days += 1.0
                        late_days += 1
                    elif is_half:
                        in_m = sample_minutes(in_min_m, in_max_m, cfg.time_distribution, self.rng)
                        out_m = max(in_m + 90, half_out_m + self.rng.randint(-10, 10))
                        status = "half_day"
                        present_days += 0.5
                        half_days += 1
                    else:
                        in_m = sample_minutes(in_min_m, in_max_m, cfg.time_distribution, self.rng)
                        out_m = sample_minutes(out_min_m, out_max_m, cfg.time_distribution, self.rng)
                        if min_shift_mins > 0 and (out_m - in_m) < min_shift_mins:
                            out_m = min(1435, max(out_m, in_m + min_shift_mins))
                        status = "present"
                        present_days += 1.0

                    if out_m <= in_m:
                        out_m = min(1435, in_m + 60)

                    dur = out_m - in_m
                    total_mins += dur
                    punches[d] = DayPunch(
                        d,
                        status,
                        format_minutes(in_m, cfg.time_format),
                        format_minutes(out_m, cfg.time_format),
                        duration_mins=dur,
                    )

            pct = round((present_days / working_days) * 100.0, 1) if working_days > 0 else 100.0
            rows.append(
                CandidateBlockResult(
                    sl_no=idx,
                    candidate=cand,
                    punches=punches,
                    working_days=working_days,
                    present_days=present_days,
                    absent_days=absent_days,
                    late_days=late_days,
                    half_days=half_days,
                    total_hours=round(total_mins / 60.0, 1),
                    attendance_pct=pct,
                    compliant=pct >= cfg.compliance_threshold_pct,
                )
            )

        return AttendanceBlock(
            sheet_name=sheet_name,
            month_label=month_label,
            span_start=span_start,
            span_end=span_end,
            dates=dates,
            holiday_map=holiday_map,
            weekend_dates=weekend_dates,
            rows=rows,
        )
`);

  // =========================================================================
  // SECTION 4: STYLED EXCEL WRITER (+ EXECUTIVE SUMMARY SHEET)
  // =========================================================================
  sections.push(`
# =============================================================================
# SECTION 4 — ENTERPRISE EXCEL WRITER (OPENPYXL)
# =============================================================================

STATUS_STYLE_MAP: Dict[str, Tuple[str, str, bool]] = {
    "weekend": ("E2E8F0", "475569", False),
    "holiday": ("DBEAFE", "1E40AF", True),
    "absent": ("FEE2E2", "991B1B", True),
    "late": ("FEF3C7", "92400E", False),
    "half_day": ("F3E8FF", "6B21A8", False),
    "inactive": ("F1F5F9", "94A3B8", False),
}


def build_output_filename(cfg: AttendanceConfig, ext: str = "xlsx") -> str:
    clean = lambda s: "".join(ch if ch.isalnum() else "_" for ch in s.strip()).strip("_")
    parts = [p for p in (clean(cfg.district), clean(cfg.location), clean(cfg.batch_no)) if p]
    s_tag = datetime.strptime(cfg.start_date, "%Y-%m-%d").strftime("%d%m%Y")
    e_tag = datetime.strptime(cfg.end_date, "%Y-%m-%d").strftime("%d%m%Y")
    prefix = f"{'_'.join(parts)}_" if parts else ""
    return f"{prefix}Attendance_{s_tag}_{e_tag}.{ext}"


def write_excel_workbook(
    blocks: Sequence[AttendanceBlock],
    cfg: AttendanceConfig,
    out_path: str | Path,
) -> Path:
    """Write single/multi-month styled attendance workbook with optional Executive Summary."""
    if not blocks:
        raise ValueError("Cannot write Excel file with zero attendance blocks.")

    theme = EXCEL_THEMES.get(cfg.excel_theme, EXCEL_THEMES["classic_navy"])
    thin = Side(border_style="thin", color=theme["border"])
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    left_mid = Alignment(horizontal="left", vertical="center", wrap_text=True)

    out_file = Path(out_path)
    sheet_cursors: Dict[str, int] = {}

    from openpyxl import Workbook
    wb = Workbook()
    # Remove default blank sheet once we create real sheets
    default_sheet = wb.active

    for blk in blocks:
        if blk.sheet_name in wb.sheetnames:
            ws = wb[blk.sheet_name]
        else:
            ws = wb.create_sheet(title=blk.sheet_name)
            sheet_cursors[blk.sheet_name] = 1
            if cfg.landscape_print:
                ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                ws.page_setup.paperSize = ws.PAPERSIZE_A4
                ws.sheet_properties.pageSetUpPr.fitToPage = True
                ws.page_setup.fitToWidth = 1
                ws.page_setup.fitToHeight = 0

        start_row = sheet_cursors.get(blk.sheet_name, 1)
        total_cols = 3 + len(blk.dates)
        last_letter = get_column_letter(total_cols)

        # 6-row merged header banner
        center_parts = [
            cfg.center_name.strip(),
            f"(Dist: {cfg.district.strip()})" if cfg.district.strip() else "",
            f"(Batch: {cfg.batch_no.strip()})" if cfg.batch_no.strip() else "",
        ]
        period_line = f"Period: {blk.span_start.strftime('%d-%m-%Y')} - {blk.span_end.strftime('%d-%m-%Y')}"
        if cfg.location.strip():
            period_line += f" | Location: {cfg.location.strip()}"
        banner_lines = [
            "BIOMETRIC ATTENDANCE SHEET",
            " ".join(p for p in center_parts if p),
            f"{cfg.report_title}{' — ' + cfg.scheme_name.strip() if cfg.scheme_name.strip() else ''}",
            blk.month_label,
            period_line,
            f"Qualification Pack: {cfg.qualification_pack.strip()}" if cfg.qualification_pack.strip() else "",
        ]

        for i, text in enumerate(banner_lines):
            r = start_row + i
            ws.merge_cells(f"A{r}:{last_letter}{r}")
            cell = ws[f"A{r}"]
            cell.value = text
            is_title = i == 0
            cell.font = Font(
                name="Calibri",
                size=14 if is_title else (11 if i == 1 else 10),
                bold=i in (0, 1, 2, 5),
                color=theme["banner_fg"] if is_title else theme["sub_fg"],
            )
            cell.alignment = center
            cell.fill = PatternFill("solid", fgColor=theme["banner_bg"] if is_title else theme["sub_bg"])
            ws.row_dimensions[r].height = 25 if is_title else 19
            for c in range(1, total_cols + 1):
                ws.cell(r, c).border = border

        r_date = start_row + 6
        r_week = start_row + 7
        r_data = start_row + 8
        ws.row_dimensions[r_date].height = 22
        ws.row_dimensions[r_week].height = 18

        hdr_fill = PatternFill("solid", fgColor=theme["hdr_bg"])
        wk_fill = PatternFill("solid", fgColor=theme["wk_bg"])
        zebra_fill = PatternFill("solid", fgColor=theme["zebra"])

        for idx_c, label in enumerate(("Sl No", "Candidate Name", "Type"), start=1):
            c_top = ws.cell(r_date, idx_c, label)
            c_top.font = Font(name="Calibri", size=10, bold=True, color=theme["hdr_fg"])
            c_top.alignment = center
            c_top.fill = hdr_fill
            c_top.border = border

            c_sub = ws.cell(r_week, idx_c, "")
            c_sub.fill = wk_fill
            c_sub.border = border

        for d_idx, dt in enumerate(blk.dates):
            col = 4 + d_idx
            c_dt = ws.cell(r_date, col, dt.strftime("%d-%m-%Y"))
            c_dt.font = Font(name="Calibri", size=9, bold=True, color=theme["hdr_fg"])
            c_dt.alignment = center
            c_dt.fill = hdr_fill
            c_dt.border = border

            is_hol = dt in blk.holiday_map
            is_wk = dt in blk.weekend_dates
            c_wk = ws.cell(r_week, col, f"{dt.strftime('%a')}*" if is_hol else dt.strftime("%a"))
            c_wk.font = Font(
                name="Calibri",
                size=9,
                bold=is_hol or is_wk,
                color="1E40AF" if is_hol else ("991B1B" if is_wk else theme["wk_fg"]),
            )
            c_wk.alignment = center
            c_wk.fill = PatternFill("solid", fgColor="DBEAFE") if is_hol else wk_fill
            c_wk.border = border

        cur_row = r_data
        for row_i, res in enumerate(blk.rows):
            r_in, r_out = cur_row, cur_row + 1
            is_zebra = row_i % 2 == 1

            ws.merge_cells(start_row=r_in, start_column=1, end_row=r_out, end_column=1)
            ws.merge_cells(start_row=r_in, start_column=2, end_row=r_out, end_column=2)

            c_sl = ws.cell(r_in, 1, str(res.sl_no))
            c_sl.font = Font(name="Calibri", size=10, bold=True)
            c_sl.alignment = center

            c_nm = ws.cell(r_in, 2, res.candidate.name)
            c_nm.font = Font(name="Calibri", size=10, bold=True)
            c_nm.alignment = left_mid

            c_tin = ws.cell(r_in, 3, "In")
            c_tin.font = Font(name="Calibri", size=9, bold=True, color="0369A1")
            c_tin.alignment = center

            c_tout = ws.cell(r_out, 3, "Out")
            c_tout.font = Font(name="Calibri", size=9, bold=True, color="475569")
            c_tout.alignment = center

            for r_x in (r_in, r_out):
                for c_x in (1, 2, 3):
                    cell_x = ws.cell(r_x, c_x)
                    cell_x.border = border
                    if is_zebra:
                        cell_x.fill = zebra_fill

            for d_idx, dt in enumerate(blk.dates):
                col = 4 + d_idx
                p = res.punches[dt]
                cin = ws.cell(r_in, col, p.in_str)
                cout = ws.cell(r_out, col, p.out_str)
                style = STATUS_STYLE_MAP.get(p.status) if cfg.color_code_cells else None
                for cell_p in (cin, cout):
                    cell_p.alignment = center
                    cell_p.border = border
                    if style:
                        cell_p.fill = PatternFill("solid", fgColor=style[0])
                        cell_p.font = Font(name="Calibri", size=9, bold=style[2], color=style[1])
                    else:
                        cell_p.font = Font(name="Calibri", size=9)
                        if is_zebra:
                            cell_p.fill = zebra_fill

            ws.row_dimensions[r_in].height = 17
            ws.row_dimensions[r_out].height = 17
            cur_row += 2

        if cfg.include_signature_block:
            sig_r = cur_row + 2
            ws.cell(sig_r, 2, f"Trainer: {cfg.trainer_name or '____________________'}").font = Font(bold=True, size=10)
            mid_c = max(5, total_cols // 2)
            ws.cell(sig_r, mid_c, f"Center Manager: {cfg.center_manager or '____________________'}").font = Font(bold=True, size=10)
            right_c = max(mid_c + 4, total_cols - 4)
            ws.cell(sig_r, right_c, "Verified By / Seal: ____________________").font = Font(bold=True, size=10)
            cur_row = sig_r + 2
        else:
            cur_row += 2

        if cfg.freeze_panes and start_row == 1:
            ws.freeze_panes = "D9"

        ws.column_dimensions["A"].width = 7
        ws.column_dimensions["B"].width = 26
        ws.column_dimensions["C"].width = 6.5
        for d_idx in range(len(blk.dates)):
            ws.column_dimensions[get_column_letter(4 + d_idx)].width = 11.5 if cfg.time_format == "12h_ampm" else 10.5

        sheet_cursors[blk.sheet_name] = cur_row + 1
${
  flags.includeStyledExcelAndSummary
    ? `
    if cfg.include_summary_sheet:
        _append_executive_summary_sheet(wb, blocks, cfg, theme, border)
`
    : ''
}
    if default_sheet.title in wb.sheetnames and len(wb.sheetnames) > 1:
        wb.remove(default_sheet)

    wb.save(out_file)
    logger.info("Saved styled Excel workbook -> %s", out_file)
    return out_file
${
  flags.includeStyledExcelAndSummary
    ? `

def _append_executive_summary_sheet(wb, blocks, cfg: AttendanceConfig, theme: Dict[str, str], border: Border) -> None:
    ws = wb.create_sheet(title="Executive_Summary")
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    center = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A1:I1")
    ws["A1"] = f"{cfg.center_name} — Consolidated Attendance Summary"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color=theme["banner_fg"])
    ws["A1"].alignment = center
    ws["A1"].fill = PatternFill("solid", fgColor=theme["banner_bg"])
    ws.row_dimensions[1].height = 28

    headers = [
        "Sl No", "Roll / ID", "Candidate Name", "Working Days",
        "Present Days", "Absent Days", "Late Punches", "Attendance %",
        f"Compliance (>={cfg.compliance_threshold_pct:g}%)",
    ]
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(3, c_idx, h)
        cell.font = Font(name="Calibri", size=10, bold=True, color=theme["hdr_fg"])
        cell.alignment = center
        cell.fill = PatternFill("solid", fgColor=theme["hdr_bg"])
        cell.border = border

    agg: Dict[str, Dict[str, float | str]] = {}
    for blk in blocks:
        for res in blk.rows:
            entry = agg.setdefault(res.candidate.name, {
                "roll": res.candidate.roll_no or "-",
                "working": 0, "present": 0.0, "absent": 0, "late": 0,
            })
            entry["working"] = int(entry["working"]) + res.working_days
            entry["present"] = float(entry["present"]) + res.present_days
            entry["absent"] = int(entry["absent"]) + res.absent_days
            entry["late"] = int(entry["late"]) + res.late_days

    for idx, (name, st) in enumerate(agg.items(), start=1):
        r = 3 + idx
        wk = int(st["working"])
        pr = float(st["present"])
        pct = round((pr / wk) * 100.0, 1) if wk > 0 else 100.0
        ok = pct >= cfg.compliance_threshold_pct
        vals = [idx, st["roll"], name, wk, pr, st["absent"], st["late"], f"{pct:.1f}%", "COMPLIANT" if ok else "SHORTFALL"]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(r, c_idx, val)
            cell.alignment = Alignment(horizontal="left" if c_idx == 3 else "center", vertical="center")
            cell.border = border
            cell.font = Font(
                name="Calibri", size=10,
                bold=c_idx in (3, 9),
                color=("166534" if ok else "991B1B") if c_idx == 9 else "0F172A",
            )
            if c_idx == 9 and cfg.color_code_cells:
                cell.fill = PatternFill("solid", fgColor="DCFCE7" if ok else "FEE2E2")

    for i, w in enumerate((8, 15, 28, 14, 14, 14, 14, 15, 20), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
`
    : ''
}`);

  // =========================================================================
  // SECTION 5: CSV & PRINTABLE HTML EXPORTERS
  // =========================================================================
  if (flags.includeCsvAndHtmlExport) {
    sections.push(`
# =============================================================================
# SECTION 5 — CSV & PRINTABLE HTML REPORT EXPORTERS
# =============================================================================

def write_csv_report(blocks: Sequence[AttendanceBlock], cfg: AttendanceConfig, out_path: str | Path) -> Path:
    """Export attendance blocks to clean RFC-4180 CSV."""
    out_file = Path(out_path)
    with out_file.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        for blk in blocks:
            writer.writerow([f"BIOMETRIC ATTENDANCE SHEET — {cfg.center_name} ({blk.month_label})"])
            hdr = ["Sl No", "Candidate Name", "Type"] + [
                f"{d.strftime('%d-%m-%Y')} ({d.strftime('%a')})" for d in blk.dates
            ]
            writer.writerow(hdr)
            for res in blk.rows:
                in_r = [res.sl_no, res.candidate.name, "In"] + [res.punches[d].in_str for d in blk.dates]
                out_r = ["", "", "Out"] + [res.punches[d].out_str for d in blk.dates]
                writer.writerow(in_r)
                writer.writerow(out_r)
            writer.writerow([])
    logger.info("Saved CSV report -> %s", out_file)
    return out_file


def write_html_report(blocks: Sequence[AttendanceBlock], cfg: AttendanceConfig, out_path: str | Path) -> Path:
    """Export clean, print-ready HTML sheet (open in browser and Save as PDF)."""
    out_file = Path(out_path)
    parts = [
        "<!doctype html><html><head><meta charset='utf-8'>",
        f"<title>{html_lib.escape(cfg.center_name)} — Attendance</title>",
        "<style>",
        "body{font-family:system-ui,sans-serif;font-size:11px;color:#0f172a;margin:20px;}",
        "table{border-collapse:collapse;width:100%;margin-bottom:28px;}",
        "th,td{border:1px solid #cbd5e1;padding:3px 5px;text-align:center;font-variant-numeric:tabular-nums;}",
        "th{background:#1e293b;color:#fff;} td.name{text-align:left;font-weight:600;}",
        ".absent{background:#fee2e2;color:#991b1b;font-weight:600;}",
        ".holiday{background:#dbeafe;color:#1e40af;font-weight:600;}",
        ".weekend{background:#e2e8f0;color:#475569;} .late{background:#fef3c7;color:#92400e;}",
        "@media print{@page{size:A4 landscape;margin:8mm;}}",
        "</style></head><body>",
    ]
    for blk in blocks:
        parts.append(f"<h2>{html_lib.escape(cfg.center_name)} — {html_lib.escape(blk.month_label)}</h2>")
        parts.append(f"<p>Batch: {html_lib.escape(cfg.batch_no)} | District: {html_lib.escape(cfg.district)} | Pack: {html_lib.escape(cfg.qualification_pack)}</p>")
        parts.append("<table><thead><tr><th>Sl</th><th>Candidate</th><th>Type</th>")
        for d in blk.dates:
            parts.append(f"<th>{d.strftime('%d/%m')}<br>{d.strftime('%a')}</th>")
        parts.append("</tr></thead><tbody>")
        for res in blk.rows:
            parts.append(f"<tr><td rowspan='2'>{res.sl_no}</td><td rowspan='2' class='name'>{html_lib.escape(res.candidate.name)}</td><td>In</td>")
            for d in blk.dates:
                p = res.punches[d]
                parts.append(f"<td class='{p.status}'>{html_lib.escape(p.in_str)}</td>")
            parts.append("</tr><tr><td>Out</td>")
            for d in blk.dates:
                p = res.punches[d]
                parts.append(f"<td class='{p.status}'>{html_lib.escape(p.out_str)}</td>")
            parts.append("</tr>")
        parts.append("</tbody></table>")
    parts.append("</body></html>")
    out_file.write_text("\\n".join(parts), encoding="utf-8")
    logger.info("Saved HTML report -> %s", out_file)
    return out_file
`);
  }

  // =========================================================================
  // SECTION 6: MODERN SCROLLABLE TABBED TKINTER GUI
  // =========================================================================
  if (flags.includeTabbedScrollableGui) {
    sections.push(`
# =============================================================================
# SECTION 6 — MODERN TABBED TKINTER GUI (WITH LIVE PREVIEW & PROFILES)
# =============================================================================

class AttendanceMakerProGUI:
    """Tabbed, resizable, high-DPI-friendly desktop GUI for Attendance Maker Pro."""

    def __init__(self, root: "tk.Tk", initial_cfg: Optional[AttendanceConfig] = None):
        self.root = root
        self.cfg = initial_cfg or AttendanceConfig()
        self.candidates: List[CandidateRecord] = []
        self.last_blocks: List[AttendanceBlock] = []
        self.selected_file_path: Optional[str] = None
        self.previous_data: Optional[PreviousAttendanceData] = None

        root.title("Biometric Attendance Maker Pro")
        root.geometry("1040x760")
        root.minsize(900, 650)

        style = ttk.Style(root)
        for theme_name in ("clam", "vista", "aqua"):
            if theme_name in style.theme_names():
                style.theme_use(theme_name)
                break

        self._build_layout()
        self._load_config_into_widgets(self.cfg)

    def _build_layout(self) -> None:
        top_bar = ttk.Frame(self.root, padding=(12, 10))
        top_bar.pack(fill="x")
        ttk.Label(top_bar, text="Biometric Attendance Maker Pro", font=("Segoe UI", 14, "bold")).pack(side="left")
${
  flags.includeJsonProfiles
    ? `        ttk.Button(top_bar, text="Load Profile", command=self._on_load_profile).pack(side="right", padx=4)
        ttk.Button(top_bar, text="Save Profile", command=self._on_save_profile).pack(side="right", padx=4)`
    : ''
}
        ttk.Button(top_bar, text="Reset Defaults", command=self._on_reset).pack(side="right", padx=4)
        ttk.Button(
            top_bar,
            text="Continue Previous Excel",
            command=self._on_continue_previous_excel,
        ).pack(side="right", padx=4)
        ttk.Button(
            top_bar,
            text="Auto-Fetch Excel (Script Folder)",
            command=self._on_auto_fetch_script_folder,
        ).pack(side="right", padx=4)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        self.tab_setup = ttk.Frame(self.notebook, padding=12)
        self.tab_roster = ttk.Frame(self.notebook, padding=12)
        self.tab_holidays = ttk.Frame(self.notebook, padding=12)
        self.tab_styling = ttk.Frame(self.notebook, padding=12)
        self.tab_preview = ttk.Frame(self.notebook, padding=12)

        self.notebook.add(self.tab_setup, text="1. Batch & Timing")
        self.notebook.add(self.tab_roster, text="2. Candidate Roster")
        self.notebook.add(self.tab_holidays, text="3. Holidays & Weekends")
        self.notebook.add(self.tab_styling, text="4. Excel & Export")
        self.notebook.add(self.tab_preview, text="5. Live Preview")

        self._build_setup_tab()
        self._build_roster_tab()
        self._build_holidays_tab()
        self._build_styling_tab()
        self._build_preview_tab()

        footer = ttk.Frame(self.root, padding=(12, 8))
        footer.pack(fill="x")
        self.status_var = tk.StringVar(value="Ready. Load a roster file or add candidates to begin.")
        ttk.Label(footer, textvariable=self.status_var, foreground="#1e3a5f").pack(side="left")
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=140)
        self.progress.pack(side="right", padx=(8, 0))
        ttk.Button(footer, text="Generate Excel (.xlsx)", command=self._on_generate_excel).pack(side="right", padx=4)
        ttk.Button(footer, text="Refresh Preview", command=self._on_refresh_preview).pack(side="right", padx=4)

    def _build_setup_tab(self) -> None:
        left = ttk.LabelFrame(self.tab_setup, text="Center & Batch Details", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=4)
        right = ttk.LabelFrame(self.tab_setup, text="Schedule, Shift & Attendance Rules", padding=10)
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=4)
        self.tab_setup.columnconfigure(0, weight=1)
        self.tab_setup.columnconfigure(1, weight=1)

        self.vars: Dict[str, tk.Variable] = {
            "center_name": tk.StringVar(),
            "district": tk.StringVar(),
            "location": tk.StringVar(),
            "batch_no": tk.StringVar(),
            "qualification_pack": tk.StringVar(),
            "scheme_name": tk.StringVar(),
            "trainer_name": tk.StringVar(),
            "center_manager": tk.StringVar(),
            "mode": tk.StringVar(value="single"),
            "in_min": tk.StringVar(),
            "in_max": tk.StringVar(),
            "out_min": tk.StringVar(),
            "out_max": tk.StringVar(),
            "time_format": tk.StringVar(value="12h_short"),
            "time_distribution": tk.StringVar(value="gaussian"),
            "auto_fill": tk.BooleanVar(value=True),
            "late_enabled": tk.BooleanVar(value=True),
            "late_chance": tk.DoubleVar(value=8.0),
            "absent_policy": tk.StringVar(value="range"),
            "min_absents": tk.IntVar(value=1),
            "max_absents": tk.IntVar(value=3),
            "compliance_pct": tk.DoubleVar(value=75.0),
            "avoid_consec": tk.BooleanVar(value=True),
            "seed": tk.StringVar(value="2026"),
        }

        fields = [
            ("Center Name:", "center_name"), ("District:", "district"),
            ("Location:", "location"), ("Batch No:", "batch_no"),
            ("Qualification Pack:", "qualification_pack"), ("Scheme:", "scheme_name"),
            ("Trainer Name:", "trainer_name"), ("Center Manager:", "center_manager"),
        ]
        for r, (lbl, key) in enumerate(fields):
            ttk.Label(left, text=lbl).grid(row=r, column=0, sticky="w", pady=4)
            ttk.Entry(left, textvariable=self.vars[key], width=34).grid(row=r, column=1, sticky="ew", padx=6, pady=4)
        left.columnconfigure(1, weight=1)

        # Right side controls
        r = 0
        ttk.Label(right, text="Generation Mode:").grid(row=r, column=0, sticky="w", pady=3)
        ttk.Combobox(
            right, textvariable=self.vars["mode"], state="readonly",
            values=["single", "months_stacked", "months_separate"], width=20,
        ).grid(row=r, column=1, sticky="w", pady=3)
        r += 1

        ttk.Label(right, text="Start / End Date:").grid(row=r, column=0, sticky="w", pady=3)
        dt_box = ttk.Frame(right)
        dt_box.grid(row=r, column=1, sticky="w", pady=3)
        self.start_picker = DateEntry(dt_box, width=11, date_pattern="yyyy-mm-dd")
        self.start_picker.pack(side="left", padx=(0, 6))
        self.end_picker = DateEntry(dt_box, width=11, date_pattern="yyyy-mm-dd")
        self.end_picker.pack(side="left")
        r += 1

        ttk.Label(right, text="Shift Preset:").grid(row=r, column=0, sticky="w", pady=3)
        self.preset_cb = ttk.Combobox(right, state="readonly", values=list(SHIFT_PRESETS.keys()), width=24)
        self.preset_cb.grid(row=r, column=1, sticky="w", pady=3)
        self.preset_cb.bind("<<ComboboxSelected>>", self._apply_shift_preset)
        r += 1

        ttk.Label(right, text="IN Range (24h HH:MM):").grid(row=r, column=0, sticky="w", pady=3)
        in_f = ttk.Frame(right)
        in_f.grid(row=r, column=1, sticky="w", pady=3)
        ttk.Entry(in_f, textvariable=self.vars["in_min"], width=8).pack(side="left")
        ttk.Label(in_f, text=" to ").pack(side="left")
        ttk.Entry(in_f, textvariable=self.vars["in_max"], width=8).pack(side="left")
        r += 1

        ttk.Label(right, text="OUT Range (24h HH:MM):").grid(row=r, column=0, sticky="w", pady=3)
        out_f = ttk.Frame(right)
        out_f.grid(row=r, column=1, sticky="w", pady=3)
        ttk.Entry(out_f, textvariable=self.vars["out_min"], width=8).pack(side="left")
        ttk.Label(out_f, text=" to ").pack(side="left")
        ttk.Entry(out_f, textvariable=self.vars["out_max"], width=8).pack(side="left")
        r += 1

        ttk.Label(right, text="Time Format / Curve:").grid(row=r, column=0, sticky="w", pady=3)
        tf_f = ttk.Frame(right)
        tf_f.grid(row=r, column=1, sticky="w", pady=3)
        ttk.Combobox(tf_f, textvariable=self.vars["time_format"], state="readonly", values=["12h_short", "12h_ampm", "24h"], width=10).pack(side="left", padx=(0, 6))
        ttk.Combobox(tf_f, textvariable=self.vars["time_distribution"], state="readonly", values=["gaussian", "uniform"], width=10).pack(side="left")
        r += 1

        ttk.Label(right, text="Absent Policy:").grid(row=r, column=0, sticky="w", pady=3)
        ttk.Combobox(
            right, textvariable=self.vars["absent_policy"], state="readonly",
            values=["none", "fixed", "range", "target_percent"], width=16,
        ).grid(row=r, column=1, sticky="w", pady=3)
        r += 1

        ttk.Label(right, text="Min / Max Absents / Mo:").grid(row=r, column=0, sticky="w", pady=3)
        ab_f = ttk.Frame(right)
        ab_f.grid(row=r, column=1, sticky="w", pady=3)
        ttk.Spinbox(ab_f, from_=0, to=20, textvariable=self.vars["min_absents"], width=5).pack(side="left")
        ttk.Label(ab_f, text=" to ").pack(side="left")
        ttk.Spinbox(ab_f, from_=0, to=20, textvariable=self.vars["max_absents"], width=5).pack(side="left")
        r += 1

        ttk.Label(right, text="Compliance Floor (%):").grid(row=r, column=0, sticky="w", pady=3)
        ttk.Spinbox(right, from_=0, to=100, textvariable=self.vars["compliance_pct"], width=7).grid(row=r, column=1, sticky="w", pady=3)
        r += 1

        ttk.Label(right, text="Random Seed (blank=auto):").grid(row=r, column=0, sticky="w", pady=3)
        ttk.Entry(right, textvariable=self.vars["seed"], width=12).grid(row=r, column=1, sticky="w", pady=3)
        r += 1

        ttk.Checkbutton(right, text="Auto-fill random IN/OUT times", variable=self.vars["auto_fill"]).grid(row=r, column=0, columnspan=2, sticky="w", pady=2)
        r += 1
        ttk.Checkbutton(right, text="Simulate occasional late punches", variable=self.vars["late_enabled"]).grid(row=r, column=0, columnspan=2, sticky="w", pady=2)
        r += 1
        ttk.Checkbutton(right, text="Avoid consecutive absent days", variable=self.vars["avoid_consec"]).grid(row=r, column=0, columnspan=2, sticky="w", pady=2)

    def _build_roster_tab(self) -> None:
        bar = ttk.Frame(self.tab_roster)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(
            bar,
            text="Auto-Fetch Excel (Script Folder)",
            command=self._on_auto_fetch_script_folder,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Import Excel / CSV File", command=self._on_import_roster).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Scan Other Folder", command=self._on_scan_folder).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Add Candidate", command=self._on_add_candidate).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Remove Selected", command=self._on_remove_candidate).pack(side="left", padx=(0, 6))
        self.roster_count_var = tk.StringVar(value="0 candidates loaded")
        ttk.Label(bar, textvariable=self.roster_count_var).pack(side="right")

        self.script_folder_info_var = tk.StringVar(
            value=f"Script folder: {get_script_directory()}"
        )
        ttk.Label(
            self.tab_roster,
            textvariable=self.script_folder_info_var,
            foreground="#475569",
        ).pack(anchor="w", pady=(0, 6))

        prev_frame = ttk.LabelFrame(
            self.tab_roster,
            text="Continue from Previous Attendance Excel (Preserve Earlier IN/OUT Timestamps)",
            padding=8,
        )
        prev_frame.pack(fill="x", pady=(0, 8))
        ttk.Button(
            prev_frame,
            text="Upload Previous Attendance Excel (.xlsx/.csv)",
            command=self._on_continue_previous_excel,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            prev_frame,
            text="Clear Previous",
            command=self._on_clear_previous_excel,
        ).pack(side="left", padx=(0, 10))
        self.prev_info_var = tk.StringVar(
            value="No previous attendance file loaded. Upload an earlier month/range Excel to preserve its exact times and extend into new dates."
        )
        ttk.Label(
            prev_frame,
            textvariable=self.prev_info_var,
            foreground="#1e3a5f",
        ).pack(side="left", fill="x", expand=True)

        self.roster_tree = ttk.Treeview(self.tab_roster, columns=("sl", "roll", "name"), show="headings", height=18)
        self.roster_tree.heading("sl", text="Sl No")
        self.roster_tree.heading("roll", text="Roll / ID")
        self.roster_tree.heading("name", text="Candidate Name")
        self.roster_tree.column("sl", width=70, anchor="center")
        self.roster_tree.column("roll", width=160, anchor="center")
        self.roster_tree.column("name", width=480, anchor="w")
        self.roster_tree.pack(fill="both", expand=True)

    def _build_holidays_tab(self) -> None:
        top = ttk.Frame(self.tab_holidays)
        top.pack(fill="x", pady=(0, 8))

        ttk.Label(top, text="Weekend Policy:").pack(side="left", padx=(0, 6))
        self.weekend_var = tk.StringVar(value=self.cfg.weekend_policy)
        ttk.Combobox(
            top, textvariable=self.weekend_var, state="readonly",
            values=["sunday", "sat_sun", "second_fourth_sat"], width=20,
        ).pack(side="left", padx=(0, 16))

        ttk.Label(top, text="Holiday Date:").pack(side="left", padx=(0, 4))
        self.hol_picker = DateEntry(top, width=11, date_pattern="yyyy-mm-dd")
        self.hol_picker.pack(side="left", padx=(0, 6))
        self.hol_label_var = tk.StringVar(value="Public Holiday")
        ttk.Entry(top, textvariable=self.hol_label_var, width=22).pack(side="left", padx=(0, 6))
        ttk.Button(top, text="Add", command=self._on_add_holiday).pack(side="left", padx=(0, 4))
        ttk.Button(top, text="Remove Selected", command=self._on_remove_holiday).pack(side="left", padx=(0, 4))
        ttk.Button(top, text="Load National Presets", command=self._on_load_presets).pack(side="right")

        self.hol_tree = ttk.Treeview(self.tab_holidays, columns=("date", "label"), show="headings", height=15)
        self.hol_tree.heading("date", text="Date (YYYY-MM-DD)")
        self.hol_tree.heading("label", text="Holiday Name")
        self.hol_tree.column("date", width=180, anchor="center")
        self.hol_tree.column("label", width=420, anchor="w")
        self.hol_tree.pack(fill="both", expand=True)

    def _build_styling_tab(self) -> None:
        box = ttk.LabelFrame(self.tab_styling, text="Excel Theme & Export Settings", padding=12)
        box.pack(fill="x", pady=4)

        self.theme_var = tk.StringVar(value=self.cfg.excel_theme)
        self.color_code_var = tk.BooleanVar(value=self.cfg.color_code_cells)
        self.sum_sheet_var = tk.BooleanVar(value=self.cfg.include_summary_sheet)
        self.freeze_var = tk.BooleanVar(value=self.cfg.freeze_panes)
        self.landscape_var = tk.BooleanVar(value=self.cfg.landscape_print)
        self.sig_var = tk.BooleanVar(value=self.cfg.include_signature_block)

        ttk.Label(box, text="Excel Color Theme:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Combobox(box, textvariable=self.theme_var, state="readonly", values=list(EXCEL_THEMES.keys()), width=20).grid(row=0, column=1, sticky="w", pady=4)

        opts = [
            ("Color-code Weekend / Holiday / Absent / Late cells", self.color_code_var),
            ("Include Consolidated Executive_Summary Worksheet", self.sum_sheet_var),
            ("Freeze Header & Name Panes (D9)", self.freeze_var),
            ("Configure Landscape A4 Fit-to-Page Print Setup", self.landscape_var),
            ("Include Trainer & Center Manager Signature Footer", self.sig_var),
        ]
        for idx, (text, var) in enumerate(opts, start=1):
            ttk.Checkbutton(box, text=text, variable=var).grid(row=idx, column=0, columnspan=2, sticky="w", pady=3)
${
  flags.includeCsvAndHtmlExport
    ? `
        alt = ttk.LabelFrame(self.tab_styling, text="Additional Export Formats", padding=12)
        alt.pack(fill="x", pady=10)
        ttk.Button(alt, text="Export CSV Report", command=self._on_export_csv).pack(side="left", padx=(0, 8))
        ttk.Button(alt, text="Export Printable HTML (PDF-ready)", command=self._on_export_html).pack(side="left")`
    : ''
}

    def _build_preview_tab(self) -> None:
        self.preview_summary_var = tk.StringVar(value="Click 'Refresh Preview' to generate a live summary.")
        ttk.Label(self.tab_preview, textvariable=self.preview_summary_var, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        cols = ("sl", "name", "working", "present", "absent", "late", "pct", "status")
        self.preview_tree = ttk.Treeview(self.tab_preview, columns=cols, show="headings", height=18)
        headers = [("sl", "Sl", 50), ("name", "Candidate", 240), ("working", "Work", 70), ("present", "Present", 75), ("absent", "Absent", 75), ("late", "Late", 65), ("pct", "Att %", 85), ("status", "Audit", 100)]
        for k, t, w in headers:
            self.preview_tree.heading(k, text=t)
            self.preview_tree.column(k, width=w, anchor="w" if k == "name" else "center")
        self.preview_tree.pack(fill="both", expand=True)

    def _apply_shift_preset(self, _event=None) -> None:
        preset = SHIFT_PRESETS.get(self.preset_cb.get())
        if preset:
            self.vars["in_min"].set(preset[0])
            self.vars["in_max"].set(preset[1])
            self.vars["out_min"].set(preset[2])
            self.vars["out_max"].set(preset[3])

    def _load_config_into_widgets(self, cfg: AttendanceConfig) -> None:
        for k in ("center_name", "district", "location", "batch_no", "qualification_pack", "scheme_name", "trainer_name", "center_manager", "mode", "in_min", "in_max", "out_min", "out_max", "time_format", "time_distribution", "absent_policy"):
            self.vars[k].set(getattr(cfg, k))
        self.start_picker.set_date(datetime.strptime(cfg.start_date, "%Y-%m-%d").date())
        self.end_picker.set_date(datetime.strptime(cfg.end_date, "%Y-%m-%d").date())
        self.vars["auto_fill"].set(cfg.auto_fill_times)
        self.vars["late_enabled"].set(cfg.late_enabled)
        self.vars["min_absents"].set(cfg.min_absents_per_month)
        self.vars["max_absents"].set(cfg.max_absents_per_month)
        self.vars["compliance_pct"].set(cfg.compliance_threshold_pct)
        self.vars["avoid_consec"].set(cfg.avoid_consecutive_absents)
        self.vars["seed"].set("" if cfg.seed is None else str(cfg.seed))
        self.weekend_var.set(cfg.weekend_policy)
        self._refresh_holiday_tree()

    def _collect_config(self) -> AttendanceConfig:
        seed_raw = str(self.vars["seed"].get()).strip()
        seed_val = int(seed_raw) if seed_raw.lstrip("-").isdigit() else None
        self.cfg.center_name = str(self.vars["center_name"].get()).strip()
        self.cfg.district = str(self.vars["district"].get()).strip()
        self.cfg.location = str(self.vars["location"].get()).strip()
        self.cfg.batch_no = str(self.vars["batch_no"].get()).strip()
        self.cfg.qualification_pack = str(self.vars["qualification_pack"].get()).strip()
        self.cfg.scheme_name = str(self.vars["scheme_name"].get()).strip()
        self.cfg.trainer_name = str(self.vars["trainer_name"].get()).strip()
        self.cfg.center_manager = str(self.vars["center_manager"].get()).strip()
        self.cfg.mode = str(self.vars["mode"].get())
        self.cfg.start_date = self.start_picker.get_date().strftime("%Y-%m-%d")
        self.cfg.end_date = self.end_picker.get_date().strftime("%Y-%m-%d")
        self.cfg.selected_months = list(range(1, 13))
        self.cfg.in_min = str(self.vars["in_min"].get()).strip()
        self.cfg.in_max = str(self.vars["in_max"].get()).strip()
        self.cfg.out_min = str(self.vars["out_min"].get()).strip()
        self.cfg.out_max = str(self.vars["out_max"].get()).strip()
        self.cfg.time_format = str(self.vars["time_format"].get())
        self.cfg.time_distribution = str(self.vars["time_distribution"].get())
        self.cfg.auto_fill_times = bool(self.vars["auto_fill"].get())
        self.cfg.late_enabled = bool(self.vars["late_enabled"].get())
        self.cfg.absent_policy = str(self.vars["absent_policy"].get())
        self.cfg.min_absents_per_month = int(self.vars["min_absents"].get())
        self.cfg.max_absents_per_month = int(self.vars["max_absents"].get())
        self.cfg.fixed_absents_per_month = int(self.vars["min_absents"].get())
        self.cfg.compliance_threshold_pct = float(self.vars["compliance_pct"].get())
        self.cfg.avoid_consecutive_absents = bool(self.vars["avoid_consec"].get())
        self.cfg.seed = seed_val
        self.cfg.weekend_policy = self.weekend_var.get()
        self.cfg.excel_theme = self.theme_var.get()
        self.cfg.color_code_cells = bool(self.color_code_var.get())
        self.cfg.include_summary_sheet = bool(self.sum_sheet_var.get())
        self.cfg.freeze_panes = bool(self.freeze_var.get())
        self.cfg.landscape_print = bool(self.landscape_var.get())
        self.cfg.include_signature_block = bool(self.sig_var.get())
        return self.cfg

    def _refresh_roster_tree(self) -> None:
        self.roster_tree.delete(*self.roster_tree.get_children())
        for idx, c in enumerate(self.candidates, start=1):
            self.roster_tree.insert("", "end", values=(idx, c.roll_no or "-", c.name))
        self.roster_count_var.set(f"{len(self.candidates)} candidates loaded")

    def _refresh_holiday_tree(self) -> None:
        self.hol_tree.delete(*self.hol_tree.get_children())
        for h in sorted(self.cfg.holidays, key=lambda x: x.date_iso):
            self.hol_tree.insert("", "end", values=(h.date_iso, h.label))

    def _on_import_roster(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.xlsm *.csv")])
        if not path:
            return
        try:
            cands, col, _ = read_roster_file(path)
            self.candidates = cands
            self.selected_file_path = path
            self._refresh_roster_tree()
            self.status_var.set(f"Loaded {len(cands)} candidates from {os.path.basename(path)} (col: {col})")
        except Exception as exc:
            messagebox.showerror("Import Error", str(exc))

    def _on_auto_fetch_script_folder(self) -> None:
        script_dir = get_script_directory()
        cands, best_file, col = auto_fetch_roster_from_folder()
        if not cands or not best_file:
            messagebox.showwarning(
                "No Excel File Found",
                f"No valid Excel or CSV candidate roster was found in:\\n{script_dir}\\n\\n"
                "Place your candidate .xlsx/.csv file next to this Python script and click Auto-Fetch again.",
            )
            return
        self.candidates = cands
        self.selected_file_path = best_file
        self._refresh_roster_tree()
        fname = os.path.basename(best_file)
        self.script_folder_info_var.set(f"Auto-fetched: {fname} (column: {col}) from {Path(best_file).parent}")
        self.status_var.set(f"Auto-fetched {len(cands)} candidates from {fname} (col: {col})")
        self.notebook.select(self.tab_roster)

    def _on_scan_folder(self) -> None:
        folder = filedialog.askdirectory()
        if not folder:
            return
        best_cands, best_file, col = auto_fetch_roster_from_folder(folder)
        if not best_cands or not best_file:
            messagebox.showwarning("No Roster Found", "No valid Excel/CSV candidate file found in folder.")
            return
        self.candidates = best_cands
        self.selected_file_path = best_file
        self._refresh_roster_tree()
        fname = os.path.basename(best_file)
        self.script_folder_info_var.set(f"Loaded: {fname} (column: {col}) from {folder}")
        self.status_var.set(f"Scanned folder -> loaded {len(best_cands)} names from {fname}")

    def _on_add_candidate(self) -> None:
        dlg = tk.Toplevel(self.root)
        dlg.title("Add Candidate")
        ttk.Label(dlg, text="Full Name:").grid(row=0, column=0, padx=8, pady=6)
        name_e = ttk.Entry(dlg, width=30)
        name_e.grid(row=0, column=1, padx=8, pady=6)
        ttk.Label(dlg, text="Roll / ID (optional):").grid(row=1, column=0, padx=8, pady=6)
        roll_e = ttk.Entry(dlg, width=30)
        roll_e.grid(row=1, column=1, padx=8, pady=6)

        def save():
            nm = name_e.get().strip()
            if nm:
                self.candidates.append(CandidateRecord(name=nm, roll_no=roll_e.get().strip()))
                self._refresh_roster_tree()
            dlg.destroy()

        ttk.Button(dlg, text="Add", command=save).grid(row=2, column=1, sticky="e", padx=8, pady=8)

    def _on_remove_candidate(self) -> None:
        selected = self.roster_tree.selection()
        if not selected:
            return
        indices = sorted([self.roster_tree.index(item) for item in selected], reverse=True)
        for i in indices:
            if 0 <= i < len(self.candidates):
                self.candidates.pop(i)
        self._refresh_roster_tree()

    def _on_add_holiday(self) -> None:
        iso = self.hol_picker.get_date().strftime("%Y-%m-%d")
        if any(h.date_iso == iso for h in self.cfg.holidays):
            messagebox.showinfo("Duplicate", f"Holiday {iso} is already in the list.")
            return
        self.cfg.holidays.append(HolidayEntry(iso, self.hol_label_var.get().strip() or "Holiday"))
        self._refresh_holiday_tree()

    def _on_remove_holiday(self) -> None:
        selected = self.hol_tree.selection()
        if not selected:
            return
        dates_to_del = {self.hol_tree.item(item, "values")[0] for item in selected}
        self.cfg.holidays = [h for h in self.cfg.holidays if h.date_iso not in dates_to_del]
        self._refresh_holiday_tree()

    def _on_load_presets(self) -> None:
        existing = {h.date_iso for h in self.cfg.holidays}
        added = 0
        for iso, lbl in INDIAN_HOLIDAY_PRESETS:
            if iso not in existing:
                self.cfg.holidays.append(HolidayEntry(iso, lbl))
                added += 1
        self._refresh_holiday_tree()
        self.status_var.set(f"Added {added} national holiday presets.")

    def _on_continue_previous_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Previous Attendance Excel to Continue From",
            filetypes=[("Attendance Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")],
        )
        if not path:
            return
        try:
            prev = parse_previous_attendance_excel(path)
        except Exception as exc:
            messagebox.showerror("Continue Previous Excel Error", str(exc))
            return

        self.previous_data = prev
        self.candidates = list(prev.candidates)
        self._refresh_roster_tree()

        # Fill any still-blank center metadata fields from the previous workbook
        for meta_key in ("center_name", "district", "batch_no", "location", "qualification_pack"):
            val = prev.detected_meta.get(meta_key, "").strip()
            if val and not str(self.vars[meta_key].get()).strip():
                self.vars[meta_key].set(val)

        if prev.min_date:
            self.start_picker.set_date(prev.min_date)
        if prev.max_date:
            cur_end = self.end_picker.get_date()
            if cur_end <= prev.max_date:
                next_yr = prev.max_date.year + (1 if prev.max_date.month == 12 else 0)
                next_mo = 1 if prev.max_date.month == 12 else prev.max_date.month + 1
                next_last_day = calendar.monthrange(next_yr, next_mo)[1]
                self.end_picker.set_date(date(next_yr, next_mo, next_last_day))

        span_txt = (
            f"{prev.min_date.strftime('%d-%m-%Y')} to {prev.max_date.strftime('%d-%m-%Y')}"
            if (prev.min_date and prev.max_date)
            else ""
        )
        self.prev_info_var.set(
            f"Loaded {prev.file_name} | {len(prev.candidates)} candidates | {prev.dates_count} dates locked ({span_txt})"
        )
        self.status_var.set(
            f"Continuing from {prev.file_name}: {prev.dates_count} previous dates locked. Set your new End Date and Generate!"
        )
        messagebox.showinfo(
            "Previous Attendance Loaded",
            f"Loaded {len(prev.candidates)} candidates and locked {prev.dates_count} dates ({span_txt}) from {prev.file_name}.\\n\\n"
            "All earlier IN/OUT times will be preserved verbatim. Adjust your End Date (e.g. extend to May/June) and click Generate Excel.",
        )

    def _on_clear_previous_excel(self) -> None:
        self.previous_data = None
        self.prev_info_var.set(
            "No previous attendance file loaded. Upload an earlier month/range Excel to preserve its exact times and extend into new dates."
        )
        self.status_var.set("Cleared previous attendance continuation lock.")

    def _on_refresh_preview(self) -> None:
        try:
            cfg = self._collect_config()
            prev_rec = self.previous_data.records if self.previous_data else None
            engine = AttendanceEngine(cfg, previous_records=prev_rec)
            self.last_blocks = engine.build_blocks(self.candidates)
            self.preview_tree.delete(*self.preview_tree.get_children())
            blk = self.last_blocks[0]
            avg = sum(r.attendance_pct for r in blk.rows) / max(1, len(blk.rows))
            ok_count = sum(1 for r in blk.rows if r.compliant)
            self.preview_summary_var.set(
                f"Previewing {blk.month_label} | {len(blk.rows)} Candidates | Avg Attendance: {avg:.1f}% | Compliant: {ok_count}/{len(blk.rows)}"
            )
            for res in blk.rows:
                self.preview_tree.insert("", "end", values=(
                    res.sl_no, res.candidate.name, res.working_days, res.present_days,
                    res.absent_days, res.late_days, f"{res.attendance_pct:.1f}%",
                    "PASS" if res.compliant else "LOW",
                ))
            self.notebook.select(self.tab_preview)
        except Exception as exc:
            messagebox.showerror("Preview Error", str(exc))

    def _on_generate_excel(self) -> None:
        try:
            cfg = self._collect_config()
            prev_rec = self.previous_data.records if self.previous_data else None
            default_name = build_output_filename(cfg, "xlsx")
            out_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", initialfile=default_name,
                filetypes=[("Excel Workbook", "*.xlsx")],
            )
            if not out_path:
                return
            self.progress.start(10)
            self.status_var.set("Generating styled Excel workbook...")

            def worker():
                try:
                    blocks = AttendanceEngine(cfg, previous_records=prev_rec).build_blocks(self.candidates)
                    saved = write_excel_workbook(blocks, cfg, out_path)
                    self.root.after(0, lambda: self._on_export_done(f"Excel saved: {saved}"))
                except Exception as exc:
                    self.root.after(0, lambda: self._on_export_err(str(exc)))

            threading.Thread(target=worker, daemon=True).start()
        except Exception as exc:
            messagebox.showerror("Configuration Error", str(exc))
${
  flags.includeCsvAndHtmlExport
    ? `
    def _on_export_csv(self) -> None:
        try:
            cfg = self._collect_config()
            prev_rec = self.previous_data.records if self.previous_data else None
            out_path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=build_output_filename(cfg, "csv"))
            if not out_path:
                return
            blocks = AttendanceEngine(cfg, previous_records=prev_rec).build_blocks(self.candidates)
            saved = write_csv_report(blocks, cfg, out_path)
            messagebox.showinfo("Exported", f"CSV saved:\\n{saved}")
        except Exception as exc:
            messagebox.showerror("CSV Export Error", str(exc))

    def _on_export_html(self) -> None:
        try:
            cfg = self._collect_config()
            prev_rec = self.previous_data.records if self.previous_data else None
            out_path = filedialog.asksaveasfilename(defaultextension=".html", initialfile=build_output_filename(cfg, "html"))
            if not out_path:
                return
            blocks = AttendanceEngine(cfg, previous_records=prev_rec).build_blocks(self.candidates)
            saved = write_html_report(blocks, cfg, out_path)
            messagebox.showinfo("Exported", f"Printable HTML saved:\\n{saved}")
        except Exception as exc:
            messagebox.showerror("HTML Export Error", str(exc))`
    : ''
}

    def _on_export_done(self, msg: str) -> None:
        self.progress.stop()
        self.status_var.set(msg)
        messagebox.showinfo("Success", msg)

    def _on_export_err(self, err: str) -> None:
        self.progress.stop()
        self.status_var.set("Generation failed.")
        messagebox.showerror("Generation Error", err)
${
  flags.includeJsonProfiles
    ? `
    def _on_save_profile(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".json", initialfile="attendance_profile.json")
        if path:
            self._collect_config().save_profile(path)
            self.status_var.set(f"Saved profile to {os.path.basename(path)}")

    def _on_load_profile(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("JSON Profile", "*.json")])
        if path:
            try:
                self.cfg = AttendanceConfig.load_profile(path)
                self._load_config_into_widgets(self.cfg)
                self.status_var.set(f"Loaded profile {os.path.basename(path)}")
            except Exception as exc:
                messagebox.showerror("Profile Error", str(exc))`
    : ''
}

    def _on_reset(self) -> None:
        self.cfg = AttendanceConfig()
        self.previous_data = None
        if hasattr(self, "prev_info_var"):
            self.prev_info_var.set(
                "No previous attendance file loaded. Upload an earlier month/range Excel to preserve its exact times and extend into new dates."
            )
        self._load_config_into_widgets(self.cfg)
        self.status_var.set("Reset all fields to defaults.")
`);
  }

  // =========================================================================
  // SECTION 7: CLI AUTOMATION & ENTRY POINT
  // =========================================================================
  sections.push(`
# =============================================================================
# SECTION 7 — CLI AUTOMATION & ENTRY POINT
# =============================================================================
${
  flags.includeCliMode
    ? `
def run_cli(args: argparse.Namespace) -> int:
    """Headless CLI runner for automated batch attendance generation."""
    cfg = ${flags.includeJsonProfiles ? 'AttendanceConfig.load_profile(args.config) if args.config else AttendanceConfig()' : 'AttendanceConfig()'}
    previous_records: Optional[Dict[str, Dict[date, Tuple[str, str]]]] = None
    prev_candidates: List[CandidateRecord] = []

    if args.continue_from:
        prev = parse_previous_attendance_excel(args.continue_from)
        previous_records = prev.records
        prev_candidates = prev.candidates
        if not args.start and prev.min_date:
            cfg.start_date = prev.min_date.isoformat()
        logger.info(
            "Continuing from %s (%d candidates, %d dates locked: %s to %s)",
            prev.file_name,
            len(prev.candidates),
            prev.dates_count,
            prev.min_date,
            prev.max_date,
        )

    if args.start:
        cfg.start_date = args.start
    if args.end:
        cfg.end_date = args.end
    if args.mode:
        cfg.mode = args.mode
    if args.seed is not None:
        cfg.seed = args.seed

    if args.input:
        candidates, col, _ = read_roster_file(args.input, column_override=args.column)
        source_file = args.input
    elif args.auto_fetch:
        candidates, source_file, col = auto_fetch_roster_from_folder(column_override=args.column)
    elif prev_candidates:
        candidates = prev_candidates
        source_file = args.continue_from
        col = "previous_attendance"
    else:
        candidates, source_file, col = auto_fetch_roster_from_folder(column_override=args.column)

    if not candidates or not source_file:
        logger.error("No candidates found in %s", args.input or args.continue_from or get_script_directory())
        return 1

    logger.info("Loaded %d candidates from %s (column=%s)", len(candidates), source_file, col)
    blocks = AttendanceEngine(cfg, previous_records=previous_records).build_blocks(candidates)

    out_xlsx = args.output or build_output_filename(cfg, "xlsx")
    write_excel_workbook(blocks, cfg, out_xlsx)
${
  flags.includeCsvAndHtmlExport
    ? `    if args.csv:
        write_csv_report(blocks, cfg, Path(out_xlsx).with_suffix(".csv"))
    if args.html:
        write_html_report(blocks, cfg, Path(out_xlsx).with_suffix(".html"))`
    : ''
}
    return 0


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Biometric Attendance Maker Pro (GUI + CLI)")
    p.add_argument("--input", "-i", help="Path to Excel/CSV candidate roster (enables headless CLI mode)")
    p.add_argument("--auto-fetch", action="store_true", help="Auto-fetch Excel/CSV roster from the script folder (headless CLI mode)")
    p.add_argument("--continue-from", help="Path to previous attendance Excel/CSV to continue from while preserving earlier timestamps")
    p.add_argument("--output", "-o", help="Output .xlsx path")
    p.add_argument("--column", help="Override candidate name column header")
    p.add_argument("--config", "-c", help="Path to JSON configuration profile")
    p.add_argument("--start", help="Start date YYYY-MM-DD")
    p.add_argument("--end", help="End date YYYY-MM-DD")
    p.add_argument("--mode", choices=["single", "months_stacked", "months_separate"], help="Sheet layout mode")
    p.add_argument("--seed", type=int, help="Deterministic random seed")
${
  flags.includeCsvAndHtmlExport
    ? `    p.add_argument("--csv", action="store_true", help="Also export CSV report")
    p.add_argument("--html", action="store_true", help="Also export printable HTML report")`
    : ''
}
    return p
`
    : ''
}

def main() -> None:
${
  flags.includeCliMode
    ? `    parser = build_arg_parser()
    args = parser.parse_args()
    if args.input or args.auto_fetch or args.continue_from:
        raise SystemExit(run_cli(args))`
    : ''
}
${
  flags.includeTabbedScrollableGui
    ? `    if not GUI_AVAILABLE:
        raise SystemExit("Tkinter/tkcalendar not available. Install tkcalendar or run with --input in CLI mode.")
    root = tk.Tk()
    AttendanceMakerProGUI(root)
    root.mainloop()`
    : `    logger.info("GUI disabled in this build. Pass --input <roster.xlsx> to generate attendance.")`
}


if __name__ == "__main__":
    main()
`);

  return sections.join('\n');
}

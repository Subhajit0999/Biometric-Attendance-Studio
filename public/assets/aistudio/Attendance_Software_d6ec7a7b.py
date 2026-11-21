# -------------------------------------------------------------
# PART 1 — IMPORTS / CONFIG / NAME READER / TIME UTILITIES
# -------------------------------------------------------------

import os
import glob
import random
import calendar
from datetime import datetime, timedelta, date
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side

# -------------------------------------------------------------
# CONFIG
# -------------------------------------------------------------
CANDIDATE_HEADER_KEYWORDS = [
    "candidate", "name", "student", "trainee", "participant",
    "candidate name", "student name"
]

# Global holiday list
HOLIDAYS = []


# -------------------------------------------------------------
# FIND NAMES EXCEL FILE
# -------------------------------------------------------------
def find_name_files(folder="."):
    exts = ("*.xlsx", "*.xls", "*.xlsm", "*.csv")
    files = []
    for e in exts:
        files.extend(glob.glob(os.path.join(folder, e)))
    files.sort()
    return files


# -------------------------------------------------------------
# FIND CANDIDATE COLUMN IN EXCEL
# -------------------------------------------------------------
def sniff_candidate_column(df):
    if df is None or df.shape[1] == 0:
        return None

    for col in df.columns:
        col_norm = str(col).strip().lower()
        for kw in CANDIDATE_HEADER_KEYWORDS:
            if kw in col_norm:
                return col

    first = str(df.columns[0]).strip().lower()
    if any(x in first for x in ["sl", "s.no", "s/no", "sl no", "slno"]):
        if df.shape[1] >= 2:
            return df.columns[1]
        return df.columns[0]

    return df.columns[0]


# -------------------------------------------------------------
# READ CANDIDATE NAMES FROM FILE
# -------------------------------------------------------------
def read_file_candidates(path):
    try:
        if path.lower().endswith(".csv"):
            df_head = pd.read_csv(path, dtype=str)
        else:
            try:
                df_head = pd.read_excel(path, dtype=str, engine="openpyxl")
            except:
                df_head = pd.read_excel(path, dtype=str)
    except:
        return [], None, None

    df_head = df_head.dropna(how="all")
    if df_head.shape[0] == 0:
        return [], None, df_head

    chosen = sniff_candidate_column(df_head)
    vals = []

    for v in df_head[chosen].astype(str).tolist():
        s = str(v).strip()
        if s == "" or s.lower() in ("nan", "none"):
            continue
        if str(chosen).strip().lower() == s.lower():
            continue
        vals.append(s)

    return vals, chosen, df_head


# -------------------------------------------------------------
# TIME PARSE + RANDOM TIME PICKER
# -------------------------------------------------------------
def parse_time_str(time_str, ampm):
    t = datetime.strptime(time_str.strip(), "%I:%M")
    if ampm.upper() == "PM" and t.hour != 12:
        t = t.replace(hour=t.hour + 12)
    if ampm.upper() == "AM" and t.hour == 12:
        t = t.replace(hour=0)
    return t


def random_time_between(min_dt, max_dt):
    if max_dt < min_dt:
        min_dt, max_dt = max_dt, min_dt
    tot = int((max_dt - min_dt).total_seconds())

    if tot <= 0:
        chosen = min_dt
    else:
        chosen = min_dt + timedelta(seconds=random.randint(0, tot))

    return chosen.strftime("%I:%M")  # always 12-hour format


# -------------------------------------------------------------
# MONTH TOOLS
# -------------------------------------------------------------
def month_first_last(year, month):
    first = date(year, month, 1)
    last_day = calendar.monthrange(year, month)[1]
    last = date(year, month, last_day)
    return first, last


def month_name(month_int):
    return calendar.month_name[month_int]


# -------------------------------------------------------------
# PART 1 END
# -------------------------------------------------------------
# -------------------------------------------------------------
# PART 2 — ATTENDANCE ENGINE (HOLIDAYS + RANDOM ABSENTS)
# -------------------------------------------------------------

def parse_holiday_list(holiday_list):
    """Convert dd-mm-yyyy / date objects into a set of date objects."""
    s = set()
    for itm in holiday_list:
        if isinstance(itm, date):
            s.add(itm)
            continue
        try:
            d = datetime.strptime(str(itm), "%d-%m-%Y").date()
            s.add(d)
        except:
            try:
                d = datetime.strptime(str(itm), "%Y-%m-%d").date()
                s.add(d)
            except:
                pass
    return s


# -------------------------------------------------------------
# MAIN ATTENDANCE GENERATOR
# -------------------------------------------------------------
def build_attendance_df_for_span(
        names_original,
        span_start: date,
        span_end: date,
        auto_fill: bool,
        in_min_dt,
        in_max_dt,
        out_min_dt,
        out_max_dt,
        holidays_for_span,
        random_absent_enabled: bool
):
    """
    Generates attendance table for a date range.
    Includes:
    - Sunday fixed 0.00
    - Holiday list 0.00
    - Optional 2 random absents/month per candidate
    """

    # ---------- CLEAN NAMES ----------
    names = [str(n).strip() for n in names_original if str(n).strip() != ""]

    # ---------- BUILD DATE RANGE ----------
    dates = []
    cur = span_start
    while cur <= span_end:
        dates.append(cur)
        cur += timedelta(days=1)

    # ---------- MONTH-WORKING DATES (no Sundays, no holidays) ----------
    month_working = {}
    for dt in dates:
        ym = (dt.year, dt.month)
        if ym not in month_working:
            month_working[ym] = []

        if dt.weekday() == 6:  # Sunday
            continue
        if dt in holidays_for_span:
            continue

        month_working[ym].append(dt)

    # ---------- BUILD ABSENT MAP ----------
    per_name_absents = {}

    if random_absent_enabled:
        # Assign exactly 2 absents per month (if possible)
        for nm in names:
            per_name_absents[nm] = {}
            for ym, wdates in month_working.items():

                if len(wdates) >= 2:
                    pick_n = 2
                else:
                    pick_n = len(wdates)

                if pick_n <= 0:
                    chosen = set()
                else:
                    try:
                        chosen = set(random.sample(wdates, pick_n))
                    except:
                        chosen = set(wdates[:pick_n])

                per_name_absents[nm][ym] = chosen
    else:
        # No absents at all
        for nm in names:
            per_name_absents[nm] = {ym: set() for ym in month_working.keys()}

    # ---------- BUILD ROWS ----------
    rows = []
    sunday_count = 0

    for nm in names:
        in_row = {"Sl No": "", "Candidate Name": nm, "Type": "In"}
        out_row = {"Sl No": "", "Candidate Name": "", "Type": "Out"}

        for dt in dates:
            key = dt.strftime("%d-%m-%Y")

            # SUNDAY
            if dt.weekday() == 6:
                in_val = "0.00"
                out_val = "0.00"
                sunday_count += 1

            # HOLIDAY
            elif dt in holidays_for_span:
                in_val = "0.00"
                out_val = "0.00"

            # ABSENTS
            else:
                ym = (dt.year, dt.month)
                absents_for_nm = per_name_absents.get(nm, {}).get(ym, set())

                if dt in absents_for_nm:
                    in_val = "0.00"
                    out_val = "0.00"

                else:
                    # PRESENT → Give IN/OUT time
                    if auto_fill:
                        in_start = in_min_dt.replace(year=dt.year, month=dt.month, day=dt.day)
                        in_end = in_max_dt.replace(year=dt.year, month=dt.month, day=dt.day)
                        out_start = out_min_dt.replace(year=dt.year, month=dt.month, day=dt.day)
                        out_end = out_max_dt.replace(year=dt.year, month=dt.month, day=dt.day)

                        in_val = random_time_between(in_start, in_end)
                        out_val = random_time_between(out_start, out_end)
                    else:
                        in_val = ""
                        out_val = ""

            in_row[key] = in_val
            out_row[key] = out_val

        rows.append(in_row)
        rows.append(out_row)

    # ---------- BUILD DATAFRAME ----------
    df = pd.DataFrame(rows)

    # Enforce column order
    cols = ["Sl No", "Candidate Name", "Type"] + [d.strftime("%d-%m-%Y") for d in dates]

    for c in cols:
        if c not in df.columns:
            df[c] = ""

    df = df[cols]

    # Renumber Sl No for "In" rows
    # NOTE: the "Sl No" column was seeded with "" (empty strings), so pandas
    # locks it to a strict string dtype. Assigning a raw int (e.g. 1) into
    # a strict-string column raises:
    #   TypeError: Invalid value '1' for dtype 'str'.
    # so we must assign the serial number as a string.
    in_indices = df.index[df["Type"].astype(str).str.lower() == "in"].tolist()
    for new_sl, idx in enumerate(in_indices, start=1):
        df.at[idx, "Sl No"] = str(new_sl)

    return df, dates, sunday_count


# -------------------------------------------------------------
# PART 2 END
# -------------------------------------------------------------
# -------------------------------------------------------------
# PART 3 — GUI (HOLIDAYS + RANDOM ABSENTS + REFRESH + PREVIEW)
# -------------------------------------------------------------

class AttendanceApp:
    def __init__(self, root):
        self.root = root
        root.title("Biometric Attendance Maker - Final Version")

        frm = ttk.Frame(root, padding=10)
        frm.grid(sticky="nsew")
        root.columnconfigure(0, weight=1)

        r = 0

        # -----------------------------------------------------
        # BASIC ENTRY FIELDS
        # -----------------------------------------------------
        ttk.Label(frm, text="Center Name:").grid(column=0, row=r, sticky="w")
        self.center_entry = ttk.Entry(frm, width=40)
        self.center_entry.grid(column=1, row=r, sticky="w", padx=6)
        r += 1

        ttk.Label(frm, text="District:").grid(column=0, row=r, sticky="w")
        self.dist_entry = ttk.Entry(frm, width=40)
        self.dist_entry.grid(column=1, row=r, sticky="w", padx=6)
        r += 1

        ttk.Label(frm, text="Location:").grid(column=0, row=r, sticky="w")
        self.location_entry = ttk.Entry(frm, width=40)
        self.location_entry.grid(column=1, row=r, sticky="w", padx=6)
        r += 1

        ttk.Label(frm, text="Batch No:").grid(column=0, row=r, sticky="w")
        self.batch_entry = ttk.Entry(frm, width=20)
        self.batch_entry.grid(column=1, row=r, sticky="w", padx=6)
        r += 1

        ttk.Label(frm, text="Qualification Pack:").grid(column=0, row=r, sticky="w")
        self.qual_entry = ttk.Entry(frm, width=40)
        self.qual_entry.insert(0, "Broadband Technician")
        self.qual_entry.grid(column=1, row=r, sticky="w", padx=6)
        r += 1

        ttk.Separator(frm).grid(column=0, columnspan=4, row=r, sticky="ew", pady=8)
        r += 1

        # -----------------------------------------------------
        # MODE SELECTION
        # -----------------------------------------------------
        ttk.Label(frm, text="Mode:").grid(column=0, row=r, sticky="w")
        self.mode_var = tk.StringVar(value="single")

        ttk.Radiobutton(frm, text="Single Range",
                        variable=self.mode_var, value="single",
                        command=self._on_mode_change).grid(column=1, row=r, sticky="w")

        ttk.Radiobutton(frm, text="By Months (Stacked)",
                        variable=self.mode_var, value="months",
                        command=self._on_mode_change).grid(column=2, row=r, sticky="w")
        r += 1

        # -----------------------------------------------------
        # DATE RANGE
        # -----------------------------------------------------
        ttk.Label(frm, text="Start Date:").grid(column=0, row=r, sticky="w")
        self.start_date = DateEntry(frm, width=12, date_pattern="dd-mm-yyyy")
        self.start_date.grid(column=1, row=r, sticky="w")

        ttk.Label(frm, text="End Date:").grid(column=2, row=r, sticky="w")
        self.end_date = DateEntry(frm, width=12, date_pattern="dd-mm-yyyy")
        self.end_date.grid(column=3, row=r, sticky="w")
        r += 1

        # -----------------------------------------------------
        # MULTI-MONTH SELECTION
        # -----------------------------------------------------
        ttk.Label(frm, text="Select Months (For 'By Months'):").grid(column=0, row=r, sticky="w")

        months = [calendar.month_name[i] for i in range(1, 13)]
        self.months_lb = tk.Listbox(frm, selectmode="multiple", exportselection=False, height=6)

        for m in months:
            self.months_lb.insert("end", m)

        self.months_lb.grid(column=1, row=r, columnspan=3, sticky="w")
        r += 1

        ttk.Separator(frm).grid(column=0, columnspan=4, row=r, sticky="ew", pady=8)
        r += 1

        # -----------------------------------------------------
        # TIME RANGE SECTION
        # -----------------------------------------------------
        ttk.Label(frm, text="IN Time Range (HH:MM):").grid(column=0, row=r, sticky="w")

        self.in_min_entry = ttk.Entry(frm, width=8)
        self.in_min_entry.insert(0, "10:00")
        self.in_min_entry.grid(column=1, row=r, sticky="w")

        self.in_min_ampm = ttk.Combobox(frm, values=["AM", "PM"], width=3)
        self.in_min_ampm.set("AM")
        self.in_min_ampm.grid(column=1, row=r, padx=(80, 0), sticky="w")

        self.in_max_entry = ttk.Entry(frm, width=8)
        self.in_max_entry.insert(0, "10:30")
        self.in_max_entry.grid(column=2, row=r, sticky="w")

        self.in_max_ampm = ttk.Combobox(frm, values=["AM", "PM"], width=3)
        self.in_max_ampm.set("AM")
        self.in_max_ampm.grid(column=2, row=r, padx=(80, 0), sticky="w")
        r += 1

        ttk.Label(frm, text="OUT Time Range (HH:MM):").grid(column=0, row=r, sticky="w")

        self.out_min_entry = ttk.Entry(frm, width=8)
        self.out_min_entry.insert(0, "04:00")
        self.out_min_entry.grid(column=1, row=r, sticky="w")

        self.out_min_ampm = ttk.Combobox(frm, values=["AM", "PM"], width=3)
        self.out_min_ampm.set("PM")
        self.out_min_ampm.grid(column=1, row=r, padx=(80, 0), sticky="w")

        self.out_max_entry = ttk.Entry(frm, width=8)
        self.out_max_entry.insert(0, "04:30")
        self.out_max_entry.grid(column=2, row=r, sticky="w")

        self.out_max_ampm = ttk.Combobox(frm, values=["AM", "PM"], width=3)
        self.out_max_ampm.set("PM")
        self.out_max_ampm.grid(column=2, row=r, padx=(80, 0), sticky="w")
        r += 1

        # -----------------------------------------------------
        # AUTO FILL RANDOM TIME
        # -----------------------------------------------------
        self.auto_fill_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm,
                        text="Auto-fill random IN/OUT times",
                        variable=self.auto_fill_var).grid(column=0, columnspan=3, row=r, sticky="w")
        r += 1

        # -----------------------------------------------------
        # RANDOM ABSENTS TOGGLE
        # -----------------------------------------------------
        self.random_absent_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm,
                        text="2 Random Absents Per Candidate Per Month",
                        variable=self.random_absent_var).grid(column=0, columnspan=3, row=r, sticky="w")
        r += 1

        # -----------------------------------------------------
        # NAMES FOLDER SECTION
        # -----------------------------------------------------
        ttk.Label(frm, text="Names Folder:").grid(column=0, row=r, sticky="w")

        self.names_folder_var = tk.StringVar(value=os.getcwd())
        self.names_folder_entry = ttk.Entry(frm, textvariable=self.names_folder_var, width=40)
        self.names_folder_entry.grid(column=1, row=r, sticky="w", columnspan=2)

        ttk.Button(frm, text="Browse", command=self.browse_folder).grid(column=3, row=r, sticky="w")
        r += 1

        ttk.Separator(frm).grid(column=0, columnspan=4, row=r, pady=8, sticky="ew")
        r += 1

        # -----------------------------------------------------
        # HOLIDAY SELECTION
        # -----------------------------------------------------
        ttk.Label(frm, text="Add Holiday (0.00) :").grid(column=0, row=r, sticky="w")
        r += 1

        self.holiday_date = DateEntry(frm, width=12, date_pattern="dd-mm-yyyy")
        self.holiday_date.grid(column=0, row=r, sticky="w")

        ttk.Button(frm, text="Add Holiday", command=self.add_holiday).grid(column=1, row=r, sticky="w")

        self.holiday_listbox = tk.Listbox(frm, height=4)
        self.holiday_listbox.grid(column=2, row=r, sticky="w", padx=10)
        r += 1

        ttk.Separator(frm).grid(column=0, columnspan=4, row=r, pady=8, sticky="ew")
        r += 1

        # -----------------------------------------------------
        # ACTION BUTTONS
        # -----------------------------------------------------
        ttk.Button(frm, text="Preview Names", command=self.preview_names).grid(column=0, row=r, sticky="w")
        ttk.Button(frm, text="Generate", command=self.generate).grid(column=1, row=r, sticky="w")
        ttk.Button(frm, text="Refresh", command=self.refresh_ui).grid(column=2, row=r, sticky="w")
        r += 1

        self.status = ttk.Label(frm, text="", foreground="green")
        self.status.grid(column=0, columnspan=4, row=r, sticky="w", pady=8)

        self._on_mode_change()

    # ---------------------------------------------------------
    # MODE CHANGE
    # ---------------------------------------------------------
    def _on_mode_change(self):
        if self.mode_var.get() == "months":
            self.months_lb.config(state="normal")
        else:
            self.months_lb.config(state="disabled")

    # ---------------------------------------------------------
    # ADD HOLIDAY
    # ---------------------------------------------------------
    def add_holiday(self):
        d = self.holiday_date.get_date()
        if d not in HOLIDAYS:
            HOLIDAYS.append(d)
            self.holiday_listbox.insert("end", d.strftime("%d-%m-%Y"))
        else:
            messagebox.showinfo("Already Exists", "This holiday is already added.")

    # ---------------------------------------------------------
    # REFRESH BUTTON — RESET EVERYTHING
    # ---------------------------------------------------------
    def refresh_ui(self):
        self.center_entry.delete(0, tk.END)
        self.dist_entry.delete(0, tk.END)
        self.location_entry.delete(0, tk.END)
        self.batch_entry.delete(0, tk.END)

        self.qual_entry.delete(0, tk.END)
        self.qual_entry.insert(0, "Broadband Technician")

        today = date.today()
        self.start_date.set_date(today)
        self.end_date.set_date(today)

        self.months_lb.selection_clear(0, tk.END)

        self.in_min_entry.delete(0, tk.END)
        self.in_min_entry.insert(0, "10:00")
        self.in_min_ampm.set("AM")

        self.in_max_entry.delete(0, tk.END)
        self.in_max_entry.insert(0, "10:30")
        self.in_max_ampm.set("AM")

        self.out_min_entry.delete(0, tk.END)
        self.out_min_entry.insert(0, "04:00")
        self.out_min_ampm.set("PM")

        self.out_max_entry.delete(0, tk.END)
        self.out_max_entry.insert(0, "04:30")
        self.out_max_ampm.set("PM")

        self.auto_fill_var.set(True)
        self.random_absent_var.set(True)

        self.names_folder_var.set(os.getcwd())

        HOLIDAYS.clear()
        self.holiday_listbox.delete(0, tk.END)

        self.mode_var.set("single")

        self.status.config(text="")

        messagebox.showinfo("Refreshed", "All fields have been reset successfully!")

    # ---------------------------------------------------------
    # SELECT NAMES FOLDER
    # ---------------------------------------------------------
    def browse_folder(self):
        d = filedialog.askdirectory()
        if d:
            self.names_folder_var.set(d)

    # ---------------------------------------------------------
    # PREVIEW NAMES
    # ---------------------------------------------------------
    def preview_names(self):
        folder = self.names_folder_var.get()
        files = find_name_files(folder)

        if not files:
            messagebox.showinfo("Error", "No Excel/CSV found in folder.")
            return

        best_f, best_names, best_col = None, [], None

        for f in files:
            names, col, df = read_file_candidates(f)
            if len(names) > len(best_names):
                best_f, best_names, best_col = f, names, col

        if not best_f:
            messagebox.showinfo("Error", "Could not read names.")
            return

        msg = f"Using File: {os.path.basename(best_f)}\nColumn: {best_col}\nTotal Names: {len(best_names)}\n\n"
        msg += "\n".join([f"{i+1}. {n}" for i, n in enumerate(best_names[:200])])

        top = tk.Toplevel(self.root)
        top.title("Names Preview")

        txt = tk.Text(top, width=100, height=30)
        txt.insert("1.0", msg)
        txt.config(state="disabled")
        txt.pack(expand=True, fill="both")

        self.status.config(text=f"Loaded names from {os.path.basename(best_f)}")

    # ---------------------------------------------------------
    # GET SELECTED MONTHS (MULTIPLE)
    # ---------------------------------------------------------
    def _gather_selected_months(self):
        sel = self.months_lb.curselection()
        return [i + 1 for i in sel]


# -------------------------------------------------------------
# PART 3 END
# -------------------------------------------------------------
# -------------------------------------------------------------
# PART 4 — GENERATE OUTPUT + WRITE TO EXCEL + MAIN
# -------------------------------------------------------------

    # ---------------------------------------------------------
    # GENERATE ATTENDANCE
    # ---------------------------------------------------------
    def generate(self):
        center = self.center_entry.get().strip() or "TSSC Bright Future"
        district = self.dist_entry.get().strip() or "District"
        location = self.location_entry.get().strip() or "Location"
        batch = self.batch_entry.get().strip() or "1"
        qual = self.qual_entry.get().strip() or "Broadband Technician"

        # ---------- DATE RANGE ----------
        try:
            overall_start = datetime.strptime(self.start_date.get(), "%d-%m-%Y").date()
            overall_end = datetime.strptime(self.end_date.get(), "%d-%m-%Y").date()
        except:
            messagebox.showerror("Error", "Invalid Start or End Date.")
            return

        if overall_end < overall_start:
            messagebox.showerror("Error", "End Date must be after Start Date.")
            return

        # ---------- TIME RANGE ----------
        try:
            in_min_t = parse_time_str(self.in_min_entry.get(), self.in_min_ampm.get())
            in_max_t = parse_time_str(self.in_max_entry.get(), self.in_max_ampm.get())
            out_min_t = parse_time_str(self.out_min_entry.get(), self.out_min_ampm.get())
            out_max_t = parse_time_str(self.out_max_entry.get(), self.out_max_ampm.get())
        except:
            messagebox.showerror("Error", "Invalid time format. Use HH:MM.")
            return

        in_min_dt = in_min_t.replace(year=overall_start.year, month=overall_start.month, day=overall_start.day)
        in_max_dt = in_max_t.replace(year=overall_start.year, month=overall_start.month, day=overall_start.day)
        out_min_dt = out_min_t.replace(year=overall_start.year, month=overall_start.month, day=overall_start.day)
        out_max_dt = out_max_t.replace(year=overall_start.year, month=overall_start.month, day=overall_start.day)

        # ---------- LOAD NAMES ----------
        folder = self.names_folder_var.get()
        files = find_name_files(folder)

        if not files:
            messagebox.showerror("Error", "No Excel/CSV name file found.")
            return

        best_f, best_names, best_col = None, [], None

        for f in files:
            names, col, df = read_file_candidates(f)
            if len(names) > len(best_names):
                best_f, best_names, best_col = f, names, col

        if not best_names:
            messagebox.showerror("Error", "No names found in selected folder.")
            return

        # ---------- GET HOLIDAYS ----------
        holidays_set = parse_holiday_list(HOLIDAYS)

        # ---------- OUTPUT FILE NAME ----------
        out_fname = (
            f"{district.replace(' ','_')}_{location.replace(' ','_')}_{batch}_"
            f"Attendance_{overall_start.strftime('%d%m%Y')}_{overall_end.strftime('%d%m%Y')}.xlsx"
        )
        out_path = os.path.join(os.getcwd(), out_fname)

        # -------------------------------------------------------------
        # PRE-BUILD ALL BLOCKS *BEFORE* OPENING THE EXCEL WRITER.
        #
        # This guarantees we never open/save an .xlsx file unless we
        # already know there is at least one sheet's worth of data to
        # write. Previously, validation (month selection / date overlap)
        # happened *after* pd.ExcelWriter(...) was already open, so an
        # empty selection led to a workbook with zero sheets, which
        # openpyxl refuses to save ("At least one sheet must be visible").
        # -------------------------------------------------------------
        blocks = []  # each item: (df_sheet, dates, header_info, sheet_name, start_row)

        if self.mode_var.get() == "single":

            df_sheet, dates, _ = build_attendance_df_for_span(
                best_names,
                overall_start,
                overall_end,
                bool(self.auto_fill_var.get()),
                in_min_dt, in_max_dt, out_min_dt, out_max_dt,
                holidays_set,
                bool(self.random_absent_var.get())
            )

            header_info = {
                "Center": center,
                "District": district,
                "Batch": batch,
                "MonthName": "",
                "ReportRange": f"{overall_start.strftime('%d-%m-%Y')} - {overall_end.strftime('%d-%m-%Y')}",
                "Qualification": qual,
            }

            blocks.append((df_sheet, dates, header_info, "Attendance_ALL", 1))

        # -------------------------------------------------
        # MULTI-MONTH MODE (STACKED)
        # -------------------------------------------------
        else:
            months_sel = self._gather_selected_months()
            if not months_sel:
                messagebox.showerror("Error", "Select months for By-Months mode.")
                return

            ym_list = []
            cur = date(overall_start.year, overall_start.month, 1)
            last = date(overall_end.year, overall_end.month, 1)

            while cur <= last:
                ym_list.append((cur.year, cur.month))
                if cur.month == 12:
                    cur = date(cur.year + 1, 1, 1)
                else:
                    cur = date(cur.year, cur.month + 1, 1)

            sheet_name = "Attendance_ALL"
            current_row = 1

            for yr, m in ym_list:
                if m not in months_sel:
                    continue

                first_m, last_m = month_first_last(yr, m)
                span_start = max(first_m, overall_start)
                span_end = min(last_m, overall_end)

                if span_end < span_start:
                    continue

                df_sheet, dates, _ = build_attendance_df_for_span(
                    best_names,
                    span_start,
                    span_end,
                    bool(self.auto_fill_var.get()),
                    in_min_dt, in_max_dt, out_min_dt, out_max_dt,
                    holidays_set,
                    bool(self.random_absent_var.get())
                )

                header_info = {
                    "Center": center,
                    "District": district,
                    "Batch": batch,
                    "MonthName": month_name(m),
                    "ReportRange": f"{span_start.strftime('%d-%m-%Y')} - {span_end.strftime('%d-%m-%Y')}",
                    "Qualification": qual,
                }

                blocks.append((df_sheet, dates, header_info, sheet_name, current_row))

                rows_written = 6 + 2 + df_sheet.shape[0]
                current_row += rows_written + 1

            # ---------- VALIDATE THERE IS SOMETHING TO WRITE ----------
            if not blocks:
                messagebox.showerror(
                    "Error",
                    "No data to write.\n\n"
                    "The selected month(s) don't fall within the Overall "
                    "Start Date / Overall End Date range. Please select "
                    "month(s) inside that date range and try again."
                )
                return

        # -------------------------------------------------------------
        # NOW ACTUALLY WRITE THE FILE (guaranteed >= 1 block/sheet)
        # -------------------------------------------------------------
        try:
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                for df_sheet, dates, header_info, sheet_name, start_row in blocks:
                    self._write_month_block(
                        writer, df_sheet, dates, header_info,
                        sheet_name, start_row
                    )

            messagebox.showinfo("Success", f"File saved:\n{out_path}")
            self.status.config(text=f"Saved: {out_path}")

        except Exception as e:
            messagebox.showerror("Write Error", str(e))


    # ---------------------------------------------------------
    # WRITE MONTH BLOCK INTO EXCEL
    # ---------------------------------------------------------
    def _write_month_block(self, writer, df, dates, header_info, sheet_name, start_row):

        df.to_excel(
            writer, sheet_name=sheet_name,
            index=False, header=False,
            startrow=start_row + 7
        )

        wb = writer.book
        ws = writer.sheets[sheet_name]

        last_col_idx = df.shape[1]
        last_col_letter = get_column_letter(last_col_idx)

        title_font = Font(bold=True, size=12)
        center_align = Alignment(horizontal="center", vertical="center")
        thin = Side(border_style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Row index references
        r1 = start_row
        r2 = start_row + 1
        r3 = start_row + 2
        r4 = start_row + 3
        r5 = start_row + 4
        r6 = start_row + 5
        r_date = start_row + 6
        r_weekday = start_row + 7
        r_data_start = start_row + 8

        # -----------------------------------------------------
        # MERGED HEADER BLOCK
        # -----------------------------------------------------
        ws.merge_cells(f"A{r1}:{last_col_letter}{r1}")
        ws[f"A{r1}"] = "BIOMETRIC ATTENDANCE SHEET"
        ws[f"A{r1}"].font = title_font
        ws[f"A{r1}"].alignment = center_align

        ws.merge_cells(f"A{r2}:{last_col_letter}{r2}")
        ws[f"A{r2}"] = f"{header_info['Center']} (Dist: {header_info['District']}) (Batch-{header_info['Batch']})"
        ws[f"A{r2}"].font = title_font
        ws[f"A{r2}"].alignment = center_align

        ws.merge_cells(f"A{r3}:{last_col_letter}{r3}")
        ws[f"A{r3}"] = "Monthly Attendance Report"
        ws[f"A{r3}"].font = title_font
        ws[f"A{r3}"].alignment = center_align

        ws.merge_cells(f"A{r4}:{last_col_letter}{r4}")
        ws[f"A{r4}"] = header_info["MonthName"]
        ws[f"A{r4}"].font = title_font
        ws[f"A{r4}"].alignment = center_align

        ws.merge_cells(f"A{r5}:{last_col_letter}{r5}")
        ws[f"A{r5}"] = header_info["ReportRange"]
        ws[f"A{r5}"].font = title_font
        ws[f"A{r5}"].alignment = center_align

        ws.merge_cells(f"A{r6}:{last_col_letter}{r6}")
        ws[f"A{r6}"] = header_info["Qualification"]
        ws[f"A{r6}"].font = title_font
        ws[f"A{r6}"].alignment = center_align

        # Border for header
        for rr in range(r1, r6 + 1):
            for cc in range(1, last_col_idx + 1):
                ws[f"{get_column_letter(cc)}{rr}"].border = border

        # -----------------------------------------------------
        # DATE HEADER ROW
        # -----------------------------------------------------
        ws[f"A{r_date}"] = "Sl No"
        ws[f"B{r_date}"] = "Candidate Name"
        ws[f"C{r_date}"] = "Type"

        for col in ("A", "B", "C"):
            ws[f"{col}{r_date}"].font = Font(bold=True)
            ws[f"{col}{r_date}"].alignment = center_align
            ws[f"{col}{r_date}"].border = border

        date_cols = [d.strftime("%d-%m-%Y") for d in dates]

        for j, c_name in enumerate(date_cols, start=4):
            col_letter = get_column_letter(j)
            ws[f"{col_letter}{r_date}"].value = c_name
            ws[f"{col_letter}{r_date}"].font = Font(bold=True)
            ws[f"{col_letter}{r_date}"].alignment = center_align
            ws[f"{col_letter}{r_date}"].border = border

        # -----------------------------------------------------
        # WEEKDAY ROW
        # -----------------------------------------------------
        for j, c_name in enumerate(date_cols, start=4):
            dt = datetime.strptime(c_name, "%d-%m-%Y")
            col_letter = get_column_letter(j)
            ws[f"{col_letter}{r_weekday}"] = dt.strftime("%a")
            ws[f"{col_letter}{r_weekday}"].alignment = center_align
            ws[f"{col_letter}{r_weekday}"].border = border

        for col in ("A", "B", "C"):
            ws[f"{col}{r_weekday}"].border = border

        # -----------------------------------------------------
        # APPLY BORDER + ALIGNMENT TO DATA
        # -----------------------------------------------------
        max_row = r_data_start + df.shape[0] - 1

        for rr in range(r_date, max_row + 1):
            for cc in range(1, last_col_idx + 1):
                cell = ws[f"{get_column_letter(cc)}{rr}"]
                cell.border = border
                cell.alignment = center_align

        # -----------------------------------------------------
        # AUTOSIZE COLUMNS
        # -----------------------------------------------------
        for i, colname in enumerate(df.columns, start=1):
            letter = get_column_letter(i)
            try:
                max_len = max(df[colname].astype(str).map(len).max(), len(colname)) + 2
                ws.column_dimensions[letter].width = min(max_len, 50)
            except:
                pass


# -------------------------------------------------------------
# MAIN PROGRAM
# -------------------------------------------------------------
def main():
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

# -------------------------------------------------------------
# PART 4 END — FULL SCRIPT COMPLETE
# -------------------------------------------------------------

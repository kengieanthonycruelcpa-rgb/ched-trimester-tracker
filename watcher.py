import time
import json
import os
import subprocess
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

EXCEL_FILE = "CHED 3-Year Trimester Application Tracker.xlsx"
JSON_FILE = "data.json"

# Every sheet in the workbook carries two title lines and a blank line
# above its column headers, so the real header sits on Excel row 4.
HEADER_SKIP = 3


def safe_str(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


def clean_progress(val, status):
    st = safe_str(status).lower()
    if st in ["done", "for review"]:
        return 100
    if pd.isna(val) or val == "" or val is None:
        return 0
    try:
        s_val = str(val).replace("%", "").strip()
        num = float(s_val)
        if 0 < num <= 1.0:
            num = num * 100
        return int(round(num))
    except:
        return 0


def find_col(df, possible_names):
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for name in possible_names:
        if name.lower() in cols_lower:
            return cols_lower[name.lower()]
    return None


def read_sheet(xls, name, fallback_index):
    """Read a sheet by name, falling back to position, with the title rows skipped."""
    sheet = name if name in xls.sheet_names else None
    if sheet is None:
        if fallback_index >= len(xls.sheet_names):
            return None
        sheet = xls.sheet_names[fallback_index]
    return pd.read_excel(xls, sheet_name=sheet, skiprows=HEADER_SKIP)


def build_json():
    if not os.path.exists(EXCEL_FILE):
        print(f"Error: '{EXCEL_FILE}' not found.")
        return

    print(f"Extracting data from '{EXCEL_FILE}'...")
    try:
        xls = pd.ExcelFile(EXCEL_FILE)

        # ---------- Tracker ----------
        df_tracker = read_sheet(xls, "Tracker", 0)

        id_col = find_col(df_tracker, ['#', 'item #', 'no', 'no.', 'id'])
        req_col = find_col(df_tracker, ['requirement', 'deliverable', 'item', 'task',
                                        'requirement / deliverable'])
        owner_col = find_col(df_tracker, ['owner', 'lead', 'owner / lead', 'responsible'])
        date_col = find_col(df_tracker, ['due date', 'due', 'deadline'])
        status_col = find_col(df_tracker, ['status'])
        prog_col = find_col(df_tracker, ['progress', '%', 'completion'])
        notes_col = find_col(df_tracker, ['notes', 'remarks', 'audit notes',
                                          'audit notes / remarks'])

        tracker_data = []
        current_area = ""
        for idx, row in df_tracker.iterrows():
            req_val = safe_str(row[req_col]) if req_col else ""

            # Area banner rows carry their text in the first column and
            # leave every other column blank. Remember them and move on.
            if not req_val:
                first = safe_str(row[id_col]) if id_col else ""
                if first.upper().startswith("AREA"):
                    current_area = first
                continue

            st_val = safe_str(row[status_col]) if status_col else "Not started"
            if not st_val:
                st_val = "Not started"

            prog_raw = row[prog_col] if prog_col else 0
            prog_val = clean_progress(prog_raw, st_val)

            raw_id = row[id_col] if id_col else None
            try:
                item_id = int(float(str(raw_id))) if pd.notna(raw_id) else (idx + 1)
            except:
                item_id = idx + 1

            due_date_str = safe_str(row[date_col])[:10] if date_col else ""

            tracker_data.append({
                "id": item_id,
                "area": current_area,
                "requirement": req_val,
                "owner": safe_str(row[owner_col]) if owner_col else "",
                "due_date": due_date_str,
                "status": st_val,
                "progress": prog_val,
                "notes": safe_str(row[notes_col]) if notes_col else ""
            })

        # ---------- Assignments ----------
        assignments_data = []
        df_assign = read_sheet(xls, "Assignments", 1)
        if df_assign is not None:
            p_col = find_col(df_assign, ['person or office', 'person / office', 'person',
                                         'office', 'name'])
            r_col = find_col(df_assign, ['role in project', 'project role', 'role'])
            o_col = find_col(df_assign, ['items owned', 'owned'])
            s_col = find_col(df_assign, ['items supporting', 'supporting'])
            for _, row in df_assign.iterrows():
                p_val = safe_str(row[p_col]) if p_col else ""
                if not p_val or p_val.isupper():   # skip the section banner rows
                    continue
                assignments_data.append({
                    "person": p_val,
                    "role": safe_str(row[r_col]) if r_col else "",
                    "items_owned": safe_str(row[o_col]) if o_col else "",
                    "items_supporting": safe_str(row[s_col]) if s_col else ""
                })

        # ---------- Faculty ----------
        faculty_data = []
        df_fac = read_sheet(xls, "Faculty", 3)
        if df_fac is not None:
            f_name = find_col(df_fac, ['faculty name', 'faculty member', 'name', 'faculty'])
            f_cred = find_col(df_fac, ['credential', 'credentials'])
            f_dept = find_col(df_fac, ['department', 'dept'])
            f_stat = find_col(df_fac, ['ft / pt', 'ft/pt', 'status'])
            f_comm = find_col(df_fac, ['committee assignment', 'committee'])
            f_supp = find_col(df_fac, ['tracker items supporting', 'supporting items', 'items'])
            for idx, row in df_fac.iterrows():
                fn_val = safe_str(row[f_name]) if f_name else ""
                if not fn_val:
                    continue
                faculty_data.append({
                    "id": idx + 1,
                    "name": fn_val,
                    "credential": safe_str(row[f_cred]) if f_cred else "",
                    "department": safe_str(row[f_dept]) if f_dept else "",
                    "ft_pt": safe_str(row[f_stat]) if f_stat else "",
                    "committee": safe_str(row[f_comm]) if f_comm else "",
                    "items_supporting": safe_str(row[f_supp]) if f_supp else ""
                })

        # ---------- My Outputs ----------
        outputs_data = []
        df_out = read_sheet(xls, "My Outputs", 4)
        if df_out is not None:
            o_date = find_col(df_out, ['date'])
            o_out = find_col(df_out, ['output produced', 'output delivered', 'output'])
            o_cap = find_col(df_out, ['capacity'])
            o_serv = find_col(df_out, ['tracker item(s) it serves', 'items served'])
            o_stat = find_col(df_out, ['status'])
            o_loc = find_col(df_out, ['where it lives', 'location'])
            for _, row in df_out.iterrows():
                out_val = safe_str(row[o_out]) if o_out else ""
                if not out_val:
                    continue
                outputs_data.append({
                    "date": safe_str(row[o_date])[:10] if o_date else "",
                    "output": out_val,
                    "capacity": safe_str(row[o_cap]) if o_cap else "",
                    "items_served": safe_str(row[o_serv]) if o_serv else "",
                    "status": safe_str(row[o_stat]) if o_stat else "",
                    "location": safe_str(row[o_loc]) if o_loc else ""
                })

        full_data = {
            "tracker": tracker_data,
            "assignments": assignments_data,
            "faculty": faculty_data,
            "outputs": outputs_data
        }

        # Never overwrite a good data.json with an empty one — if the sheet
        # layout changes again, keep the last known-good file and shout.
        if not tracker_data:
            print("ABORTED: no tracker rows parsed. data.json left untouched.")
            return False

        with open(JSON_FILE, "w") as f:
            json.dump(full_data, f, indent=2)
        print(f"SUCCESS: Generated '{JSON_FILE}' — "
              f"{len(tracker_data)} tracker items, {len(assignments_data)} assignments, "
              f"{len(faculty_data)} faculty, {len(outputs_data)} outputs.")
        return True

    except Exception as e:
        print(f"ERROR processing Excel file: {e}")
        return False


def sync_to_git():
    if not build_json():
        return
    try:
        subprocess.run(["git", "add", JSON_FILE], check=True)
        subprocess.run(["git", "commit", "-m", "Auto-update tracker data from Excel"], check=True)
        subprocess.run(["git", "push", "origin", "main"], check=True)
        print("Pushed data.json to GitHub successfully.")
    except Exception as e:
        print(f"Git push log: {e}")


class ExcelHandler(FileSystemEventHandler):
    def __init__(self):
        self.last_modified = 0

    def on_modified(self, event):
        if event.src_path.endswith(EXCEL_FILE):
            now = time.time()
            if now - self.last_modified > 3:
                self.last_modified = now
                print("Excel update detected. Regenerating JSON...")
                time.sleep(1)
                sync_to_git()


if __name__ == "__main__":
    sync_to_git()
    print(f"Monitoring '{EXCEL_FILE}' for changes... Press Ctrl+C to stop.")

    event_handler = ExcelHandler()
    observer = Observer()
    observer.schedule(event_handler, path=".", recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

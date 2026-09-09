import time
import subprocess
import json
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

EXCEL_FILE = "CHED 3-Year Trimester Application Tracker.xlsx"
JSON_OUTPUT = "data.json"
DEBOUNCE_INTERVAL = 4

def build_json():
    xls = pd.ExcelFile(EXCEL_FILE)
    payload = {}

    # 1. Parse Tracker Sheet
    df_tr = pd.read_excel(xls, sheet_name='Tracker', skiprows=3)
    tracker = []
    current_area = "AREA A — GOVERNANCE & SUBMISSION"
    for idx, row in df_tr.iterrows():
        item_id = str(row['#']).strip() if pd.notna(row['#']) else ''
        if item_id.startswith('AREA'):
            current_area = item_id
            continue
        if pd.isna(row['Requirement']):
            continue
        due_date = row['Due date'].strftime('%Y-%m-%d') if pd.notna(row['Due date']) and hasattr(row['Due date'], 'strftime') else ''
        try:
            pct = float(row['%']) if pd.notna(row['%']) else 0.0
        except Exception:
            pct = 0.0
            
        tracker.append({
            "area": current_area,
            "id": item_id,
            "requirement": str(row['Requirement']).strip(),
            "ched_ref": str(row['CHED reference']).strip() if pd.notna(row['CHED reference']) else '',
            "owner": str(row['Owner']).strip() if pd.notna(row['Owner']) else '',
            "support": str(row['Support']).strip() if pd.notna(row['Support']) else '',
            "due_date": due_date,
            "status": str(row['Status']).strip() if pd.notna(row['Status']) else 'Not started',
            "progress": round(pct * 100),
            "notes": str(row['Notes']).strip() if pd.notna(row['Notes']) and str(row['Notes']) != 'nan' else ''
        })
    payload['tracker'] = tracker

    # 2. Parse Assignments Sheet
    df_as = pd.read_excel(xls, sheet_name='Assignments', skiprows=3)
    assignments = []
    current_cat = "Leadership & Offices"
    for idx, row in df_as.iterrows():
        person = str(row['Person or Office']).strip() if pd.notna(row['Person or Office']) else ''
        if not person or person == 'nan':
            continue
        if pd.isna(row['Role in project']) and (person.isupper() or 'LEADS' in person or 'CHAIRS' in person or 'POOL' in person):
            current_cat = person
            continue
        assignments.append({
            "category": current_cat,
            "person": person,
            "role": str(row['Role in project']).strip() if pd.notna(row['Role in project']) and str(row['Role in project']) != 'nan' else '',
            "items_owned": str(row['Items owned']).strip() if pd.notna(row['Items owned']) and str(row['Items owned']) != 'nan' else '',
            "items_supporting": str(row['Items supporting']).strip() if pd.notna(row['Items supporting']) and str(row['Items supporting']) != 'nan' else ''
        })
    payload['assignments'] = assignments

    # 3. Parse Faculty Sheet
    df_fac = pd.read_excel(xls, sheet_name='Faculty', skiprows=2)
    df_fac.columns = [str(c).strip() for c in df_fac.iloc[0]]
    df_fac = df_fac.iloc[1:]
    faculty = []
    for idx, row in df_fac.iterrows():
        name = str(row['Faculty name']).strip() if pd.notna(row['Faculty name']) else ''
        if not name or name in ['nan', 'Faculty name']:
            continue
        faculty.append({
            "id": str(row['#']).strip() if pd.notna(row['#']) and str(row['#']) != 'nan' else '',
            "name": name,
            "credential": str(row['Credential']).strip() if pd.notna(row['Credential']) and str(row['Credential']) != 'nan' else '',
            "department": str(row['Department']).strip() if pd.notna(row['Department']) and str(row['Department']) != 'nan' else '',
            "ft_pt": str(row['FT / PT']).strip() if pd.notna(row['FT / PT']) and str(row['FT / PT']) != 'nan' else '',
            "committee": str(row['Committee assignment']).strip() if pd.notna(row['Committee assignment']) and str(row['Committee assignment']) != 'nan' else '',
            "items_supporting": str(row['Tracker items supporting']).strip() if pd.notna(row['Tracker items supporting']) and str(row['Tracker items supporting']) != 'nan' else ''
        })
    payload['faculty'] = faculty

    # 4. Parse My Outputs Sheet
    df_out = pd.read_excel(xls, sheet_name='My Outputs', skiprows=2)
    df_out.columns = [str(c).strip() for c in df_out.iloc[0]]
    df_out = df_out.iloc[1:]
    outputs = []
    for idx, row in df_out.iterrows():
        out_desc = str(row['Output produced']).strip() if pd.notna(row['Output produced']) else ''
        if not out_desc or out_desc in ['nan', 'Output produced']:
            continue
        date_str = row['Date'].strftime('%Y-%m-%d') if pd.notna(row['Date']) and hasattr(row['Date'], 'strftime') else ''
        outputs.append({
            "date": date_str,
            "output": out_desc,
            "capacity": str(row['Capacity']).strip() if pd.notna(row['Capacity']) and str(row['Capacity']) != 'nan' else '',
            "items_served": str(row['Tracker item(s) it serves']).strip() if pd.notna(row['Tracker item(s) it serves']) and str(row['Tracker item(s) it serves']) != 'nan' else '',
            "status": str(row['Status']).strip() if pd.notna(row['Status']) and str(row['Status']) != 'nan' else '',
            "location": str(row['Where it lives']).strip() if pd.notna(row['Where it lives']) and str(row['Where it lives']) != 'nan' else ''
        })
    payload['outputs'] = outputs

    with open(JSON_OUTPUT, 'w', encoding='utf-8') as f:
        json.dump(payload, f, indent=2)

class SyncHandler(FileSystemEventHandler):
    def __init__(self):
        self.last_triggered = 0

    def on_modified(self, event):
        if event.src_path.endswith(EXCEL_FILE) and "~$" not in event.src_path:
            current_time = time.time()
            if current_time - self.last_triggered > DEBOUNCE_INTERVAL:
                self.last_triggered = current_time
                print("Excel update detected. Extracting data...")
                try:
                    build_json()
                    subprocess.run(["git", "add", JSON_OUTPUT], check=True)
                    subprocess.run(["git", "commit", "-m", "Auto-update CHED tracker data"], check=True)
                    subprocess.run(["git", "push", "origin", "main"], check=True)
                    print("Updated data.json and pushed to GitHub Pages successfully.")
                except Exception as e:
                    print(f"Error during sync: {e}")

if __name__ == "__main__":
    build_json()
    event_handler = SyncHandler()
    observer = Observer()
    observer.schedule(event_handler, path=".", recursive=False)
    observer.start()
    print(f"Monitoring '{EXCEL_FILE}' for saves...")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

# Enhanced Kanji Database Tool for Google Colab

!pip install xlsxwriter --quiet

import pandas as pd
import re
import io
import os
import sqlite3
from google.colab import files
import ipywidgets as widgets
from IPython.display import display, clear_output

# --- Utility Functions ---

def extract_kanji(text):
    if pd.isna(text):
        return []
    # CJK ideographs only, no hiragana, no katakana
    return re.findall(r'[\u4E00-\u9FFF]', str(text))
    # If need hangul use  this instead:
    # re.findall(r'[\u4E00-\u9FFF\uAC00-\uD7AF]', str(text))

def read_uploaded_file(expected_format="excel_or_sqlite"):
    uploaded = files.upload()
    filename = next(iter(uploaded))

    # Validate file type
    if expected_format == "excel_or_sqlite":
        if not filename.endswith((".xlsx", ".xls", ".csv", ".sqlite")):
            raise ValueError("❌ Please upload a valid Excel (.xlsx, .xls, .csv) or SQLite (.sqlite) file.")
    elif expected_format == "excel":
        if not filename.endswith((".xlsx", ".xls", ".csv")):
            raise ValueError("❌ Please upload a valid Excel or CSV file.")

    return filename, uploaded[filename]

def make_safe_filename(raw_name, prefix="kanji_database_"):
    base = os.path.splitext(raw_name)[0]
    base = base.replace(",", "").replace(" ", "_")
    return f"{prefix}{base}.xlsx"

def write_kanji_to_excel(kanji_set, filename):
    df = pd.DataFrame(sorted(kanji_set), columns=["Kanji"])
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="KanjiDatabase")
    files.download(filename)

def write_kanji_to_sqlite(kanji_set, db_filename):
    db_name = db_filename.replace(".xlsx", ".sqlite")
    conn = sqlite3.connect(db_name)
    c = conn.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS kanji (char TEXT PRIMARY KEY)")
    for char in sorted(kanji_set):
        c.execute("INSERT OR IGNORE INTO kanji (char) VALUES (?)", (char,))
    conn.commit()
    conn.close()
    files.download(db_name)

def is_japanese_column(col_name):
    col_name = str(col_name).strip().lower()
    keywords = ["japanese", "jap", "ja", "jp", "日本", "日本語"]
    return any(k in col_name for k in keywords)

def scan_excel_for_kanji_column(file_bytes, filename, progress_bar=None):
    all_kanji = set()
    try:
        if filename.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(file_bytes))
            sheets = {"Sheet1": df}
        else:
            xls = pd.ExcelFile(io.BytesIO(file_bytes))
            sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    except Exception as e:
        raise ValueError(f"❌ Could not read the file: {e}")

    total = len(sheets)
    for i, (sheet_name, df) in enumerate(sheets.items(), start=1):
        col = next((c for c in df.columns if is_japanese_column(c)), None)
        if col:
            for val in df[col]:
                chars = extract_kanji(val)
                all_kanji.update(chars)
        if progress_bar:
            progress_bar.value = int((i / total) * 100)
    return all_kanji

# --- Main Functionalities ---

def construct_kanji_database():
    clear_output()
    print("📥 Upload Excel/CSV file with a JAPANESE column...")
    filename, file_bytes = read_uploaded_file("excel")
    progress = widgets.IntProgress(value=0, min=0, max=100)
    display(progress)
    kanji_set = scan_excel_for_kanji_column(file_bytes, filename, progress)
    safe_filename = make_safe_filename(filename)
    write_kanji_to_excel(kanji_set, safe_filename)
    write_kanji_to_sqlite(kanji_set, safe_filename)
    print("✅ Kanji database created!")

def append_to_database():
    clear_output()
    print("📥 Upload existing Kanji database (Excel/CSV)...")
    db_filename, db_bytes = read_uploaded_file("excel")
    db_df = pd.read_excel(io.BytesIO(db_bytes))
    existing_kanji = set(db_df["Kanji"].dropna().astype(str))

    print("📥 Upload Excel/CSV with NEW content...")
    update_filename, update_bytes = read_uploaded_file("excel")
    progress = widgets.IntProgress(value=0, min=0, max=100)
    display(progress)
    new_kanji = scan_excel_for_kanji_column(update_bytes, update_filename, progress)

    added_kanji = new_kanji - existing_kanji
    all_kanji = existing_kanji.union(added_kanji)

    updated_filename = make_safe_filename(db_filename, prefix="updated_")
    write_kanji_to_excel(all_kanji, updated_filename)
    write_kanji_to_sqlite(all_kanji, updated_filename)
    print("✅ Updated Kanji database created!")

def check_against_database():
    clear_output()
    print("📥 Upload Kanji database (Excel/SQLite)...")
    db_filename, db_bytes = read_uploaded_file("excel_or_sqlite")

    if db_filename.endswith(".sqlite"):
        conn = sqlite3.connect(f"/content/{db_filename}")
        df = pd.read_sql_query("SELECT char as Kanji FROM kanji", conn)
        conn.close()
    else:
        df = pd.read_excel(io.BytesIO(db_bytes))

    existing_kanji = set(df["Kanji"].dropna().astype(str))

    print("📥 Upload Excel/CSV to check against DB...")
    check_filename, check_bytes = read_uploaded_file("excel")
    xls = pd.ExcelFile(io.BytesIO(check_bytes))

    missing_report = []
    progress = widgets.IntProgress(value=0, min=0, max=100)
    display(progress)

    for i, sheet_name in enumerate(xls.sheet_names):
        df = xls.parse(sheet_name)
        col = next((c for c in df.columns if is_japanese_column(c)), None)
        if col:
            for idx, val in enumerate(df[col]):
                if pd.isna(val): continue
                full_text = str(val)
                for char in extract_kanji(val):
                    if char not in existing_kanji:
                        missing_report.append((sheet_name, idx + 2, full_text, char))
        progress.value = int((i + 1) / len(xls.sheet_names) * 100)

    if not missing_report:
        print("✅ All Kanji characters are in the database!")
    else:
        report_df = pd.DataFrame(missing_report, columns=["Sheet", "Row", "Cell Content", "Missing Character"])
        report_filename = make_safe_filename(check_filename, prefix="missing_kanji_report_")
        with pd.ExcelWriter(report_filename, engine="xlsxwriter") as writer:
            report_df.to_excel(writer, index=False, sheet_name="MissingKanji")
            worksheet = writer.sheets["MissingKanji"]
            worksheet.set_column("C:C", 50)
            worksheet.set_column("D:D", 50)
        files.download(report_filename)
        print(f"⚠️ Missing characters found. Report saved as {report_filename}")

# --- UI Menu ---

def start_menu():
    clear_output()
    def on_button_click(choice):
        button1.disabled = True
        button2.disabled = True
        button3.disabled = True
        clear_output(wait=True)
        if choice == "1":
            construct_kanji_database()
        elif choice == "2":
            append_to_database()
        elif choice == "3":
            check_against_database()

    button1 = widgets.Button(description="📥 Construct Kanji Database", layout=widgets.Layout(width='50%'))
    button2 = widgets.Button(description="📥 Append Kanji to Database", layout=widgets.Layout(width='50%'))
    button3 = widgets.Button(description="📤 Check Against Database", layout=widgets.Layout(width='50%'))

    button1.on_click(lambda b: on_button_click("1"))
    button2.on_click(lambda b: on_button_click("2"))
    button3.on_click(lambda b: on_button_click("3"))

    display(widgets.VBox([
        widgets.Label("👋 Welcome! Please choose an option:"),
        button1,
        button2,
        button3
    ]))

start_menu()

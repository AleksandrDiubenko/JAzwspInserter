import re
from google.colab import files
import io
from openpyxl import load_workbook
import os

# --- Upload Excel file ---
uploaded = files.upload()
filename = list(uploaded.keys())[0]

# --- Load workbook preserving formatting ---
wb = load_workbook(io.BytesIO(uploaded[filename]))
target_headers = {"ja", "jp", "jap", "japanese"}

# --- Your regex (mostly unchanged, only minor punctuation normalization) ---
pattern = re.compile(r"""
(
    ([一-龯]{2}|[゠-ヿ]{2,12}|こと|ところ|[一-龯](?:[ぁ-ゖ゛-ゟー](?!で))+[一-龯])
    (が|は|の|する(?!な)|から|まで|に(?![はも])|に[はも]|へ|で(?![はもす])|で[はも]|じて)
    |
    [、。？！]
    |
    (――)|(……)
    |
    [一-龯]すぎ[^たるだ]
    |
    について|に関して|ったり|とにかく|でも|[くぐ]らい|まるで|って(?![るか])|すなわち|
    [うくすつぬふむる]の[にはもが]|を|んな[のに]|ったら|として|つまり|ちょっと|ちょうど|
    だと|だけ|とは|のほうが|ないほうが|の方が|ない方が|風に|[いきしちにひみり]たくて|
    ほとんど|らしくて|らしく(?!て)
)
""", re.VERBOSE)

# --- Functions ---

def cleanup_zwsp_spacing(text):
    """Remove extra ZWSPs that are within 1–2 chars of another, preserving text."""
    if not isinstance(text, str):
        return text
    return re.sub(r'\u200B(.{1,2})\u200B', lambda m: m.group(1) + '\u200B', text)

def postprocess_ellipses(text):
    """Handle special rules for ellipses (… and ……): 
       - No ZWSP if text starts with ellipsis
       - Add ZWSP after single '…' (not '……') when mid-sentence
    """
    if not isinstance(text, str):
        return text

    # 1️⃣ Remove ZWSP immediately after starting ellipses
    text = re.sub(r'^(…{1,2})\u200B', r'\1', text)

    # 2️⃣ Add ZWSP after single ellipsis not followed by another ellipsis
    text = re.sub(r'(?<!…)(…)(?!…)(?=\S)', lambda m: m.group(1) + '\u200B', text)

    return text

def insert_zero_width_spaces(text):
    """Insert ZWSPs after pattern matches, but skip if followed only by punctuation or at end."""
    if not isinstance(text, str):
        return text

    def replacer(m):
        end = m.end()
        remainder = text[end:]
        next_char = remainder[:1]

        # 1️⃣ Skip if next char is punctuation
        if re.match(r'[、。？！,．,.!?"”」』）)]', next_char):
            return m.group(0)

        # 2️⃣ Skip if remainder (after match) consists of only punctuation or whitespace
        if re.match(r'^[、。？！…‥！？\s]*$', remainder):
            return m.group(0)

        # ✅ Otherwise, add ZWSP
        return m.group(0) + '\u200B'

    # Step 1: conditional insertion
    processed = pattern.sub(replacer, text)

    # Step 2: cleanup of near-duplicate ZWSPs
    processed = cleanup_zwsp_spacing(processed)

    # Step 3: handle ellipsis rules (no leading ZWSP, etc.)
    processed = postprocess_ellipses(processed)

    return processed

# --- Process all sheets ---
for ws in wb.worksheets:
    headers = {cell.value: cell.column for cell in ws[1] if cell.value}
    for header, col in headers.items():
        if str(header).strip().lower() in target_headers:
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col)
                if isinstance(cell.value, str):
                    new_val = insert_zero_width_spaces(cell.value)
                    if new_val != cell.value:
                        cell.value = new_val

# --- Build dynamic output filename ---
name, ext = os.path.splitext(filename)
output_filename = f"zwsp_added_{name}{ext}"

# --- Save updated workbook ---
wb.save(output_filename)
files.download(output_filename)

print(f"✅ Done! File saved as: {output_filename}")
print("✅ All formatting preserved and zero-width spaces added safely.")

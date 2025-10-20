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

# --- Main regex ---
pattern = re.compile(r"""
(
    ([一-龯]{1,2}|[゠-ヿ]{2,12}|こと|ところ|[一-龯](?:[ぁ-ゖ゛-ゟ](?!で))+[一-龯]|[゠-ヿ]{2,12}[一-龯]|もの|入り|」|たち|ここ|そこ|[一-龯]ら|(?P<double>[ぁ-ゖ゛-ゟ]{2})(?P=double)|[えけげせぜてでねめれ]る|まま|[あこそ]いつ)
    (が(?!して)|か(?![はもらえけげせぜてでねめれいきぎしちにんをうくぐすつぬむる])|か[は]|は(?!ず)|も|の(?![みにがはた為])|なく(?!て)|な(?![くのんらい])|する(?!な)|から(?!して)|まで|に(?!([はも]|ついて|関して))|
    に[はも]|へ|で(?![はもすき])|で[はも]|じて|や|との|と(?![のな])|して[はも]|して(?!はも)|ならば|なら(?![ばで]))
    |
    [、。？！・：；]
    |
    (――)|(……)|(\.\.\.)
    |
    [一-龯]すぎ[^たるだ]
    |
    について(?![はも])|について[はも]|に関して(?![はも])|に関して[はも]|[っいきぎしちにん][ただ]り|とにかく|でも|[くぐ]らい(?!は)|[くぐ]らいは|まるで|って(?![るたかも])|っても|
    すなわち|[うくぐすつぬふむる]の[にはもが]|を|んな[のに]|ったら|として|つまり|ちょっと|ちょうど|々な|々に(?![もは])|々に[もは]|(?=([一-龯]|[゠-ヿ]{2}))|
    だと|だけ|とは|[のただ]ほうが|ないほうが|[のただ]方が|ない方が|風に|[いきしちにひみり]たくて|[うくすつぬふむる]まて|[^一-龯]続く|(?=([一-龯]|[゠-ヿ]{2}))|
    ほとんど|らしくて(?!は)|らしく(?!て)|ため([にの](?![はも])|ならば|なら(?!ば))|ため[にの][はも]|為に(?![はも])|為に[はも]|わけ(では|じゃ(?!あ))|まったく|
    いきなり|すれば|れば(?=[い良善好]い)|て(?=い?ました)|しっかり|して(?=あげ([るた]|(ます|まし)))|て(?=(ください|下さい))|これまでに(?!は)|
    より(?=ずっと)|はじめて|て(?=くれ)|くなって(?!は)|され[るた](?![んの])|かった(?![んのり])|もなくて(?!は)|あらゆる|すべて(の|を|では|じゃ(?!あ))|
    もなく(?!て)|ながら|がてら|った(?![らんのり])|よりも|かも(?=[しれ])|とともに(?![はも])|と共に(?![はも])|もっとも|すべて(?!でのを)|ただの|
    どうして|どうやって|[一-龯]{2}した(?=([一-龯]{2}|こと|とこ))|のもとに|[うくすつぬふむるじの]よう[にな]|れて(?=(いき?ま|いる|いた|いな))|
    どうなるか(?!は)|どうなるかは|しばらく|[えけげせぜてでねめれ]なく(?!て)|[えけげせぜてでねめれあかさたなまら]ずに|[えけげせぜてでねめれいきしじちにみりっ]て(?=い(る|ま|く|け))|
    [一-龯]し?い(?=([一-龯]|[゠-ヿ]{2}))|[一-龯]しく(?=([一-龯]|[゠-ヿ]{2}))|べきじゃ(?!あ)|かなり(?=([一-龯]|[゠-ヿ]{2}))|[えけげせぜてでねめれ]ば(?=([一-龯]|[゠-ヿ]{2}))|
    ゆっくり(?=([一-龯]|[゠-ヿ]{2}))|ちゃんと(?=([一-龯]|[゠-ヿ]{2}))|(なければ|なきゃ)(?=(なら|いけ))|[ぁ-ゖ゛-ゟ](?=(はず|べき)だ)|[ぁ-ゖ゛-ゟ](?=[゠-ヿ]{2})|て(?=ありがと)|
    なら(?=([一-龯]|[゠-ヿ]{2}))|なのは|[えけげせぜてでねめれ]る(?=([一-龯]|[゠-ヿ]{2}))|たく(?=な[いか])|[わかさたなまら]れ[るた](?=([一-龯]|[゠-ヿ]{2}))|いくつか|[一-龯]ても|して(?=[一-龯]{2})
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
    text = re.sub(r'^(…{1,4})\u200B', r'\1', text)

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

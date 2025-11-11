import regex as re
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
    ([一-龯]{1,2}|[゠-ヿ]{2,12}|こと|ところ|[一-龯](?:[ぁ-ゖ゛-ゟ](?!で))+[一-龯]|[゠-ヿ]{2,12}[一-龯]|もの|入り|」|たち|ここ|そこ|[一-龯]ら|(?P<double>[ぁ-ゖ゛-ゟ]{2})(?P=double)|[えけげせぜてでねめれ]る|まま|[あこそ]いつ|あ[なん]た|さん|まみれ)
    (が(?!(して|った))|か(?!([はもらえけげせぜてでねめれいきぎしちにんをうくぐすつぬむるっ]|った|さ))|か[は]|は(?!ず)|も(?!の)|の(?![みにがはた為])|なく(?!て)|な(?![くのんらるい])|する(?!な)|から(?!して)|まで|に(?!([はも]|ついて|関して|すら))|
    に[はも]|へ[の]|へ(?![の])|で(?![はもすしきの])|で[はも]|じて(?!る)|や(?![かり])|と[のはか]|と(?!([のなはか]|[い言云]う))|して[はも]|して(?![はも])|ならば|なら(?![ばで]))
    |
    [、。？！・：；]
    |
    (――)|(……)|(\.\.\.)
    |
    [一-龯]すぎ[^たるだ]
    |
    について(?![はも])|について[はも]|に関して(?![はも])|に関して[はも]|[っいきぎしちにん][ただ]り|とにかく|でも|[くぐ]らい(?!は)|[くぐ]らいは|まるで|って(?![るたかも])|っても|
    すなわち|[うくぐすつぬふむる]の[にはもが]|を|んな[のに]|[って]たら|として|つまり|ちょっと|ちょうど|々な|々に(?![もは])|々に[もは]|たい(?=[一-龯])|けど|よう[なに](?=([一-龯]{2}|[゠-ヿ]{2}))|
    だと(?!は)|だとは|とは|[のただ]ほうが|ないほうが|[のただ]方が|ない方が|風に|[いきしちにひみり]たくて|[うくすつぬふむる]まて|[^一-龯]続く|ないと(?=いけ)|く(?=([一-龯]|[゠-ヿ]{2}))|
    ほとんど|らしくて(?!は)|らしく(?!て)|ため([にの](?![はも])|ならば|なら(?!ば))|ため[にの][はも]|為に(?![はも])|為に[はも]|わけ(では|じゃ(?!あ))|ほうが(?=([一-龯]|[゠-ヿ]{2}))|
    いきなり|すれば|(れば|ないと)(?=([い良善好]い|[よ良善好]か))|て(?=い?ました)|しっかり|して(?=あげ([るた]|(ます|まし)))|て(?=(ください|下さい|ちょうだい))|これまでに(?!は)|
    より(?=ずっと)|はじめて|て(?=くれ)|くなって(?!は)|され[るた](?![んの])|かった(?![んのりわっがぞぜ])|もなくて(?!は)|あらゆる|すべて(の|を|では|じゃ(?!あ))|すぐに[はも]|すぐに(?![はも])|
    もなく(?!て)|ながら|がてら|った(?![らんのりわっがぞぜ])|よりも|かも(?=[しれ])|とともに(?![はも])|と共に(?![はも])|もっとも|すべて(?!でのを)|ただの|まま(?=([一-龯]|[゠-ヿ]{2}))|
    どうして|どうやって|した(?=([一-龯]{2}|こと|とこ))|のもとに|[うくすつぬふむるじの]よう[にな]|れて(?=(いき?ま|いる|いた|いな))|じゃ(?=な[いか])|では(?=な[いか])|またしても|
    どうなるか(?!は)|どうなるかは|しばらく|[えけげせぜてでねめれ]なく(?!て)|[えけげせぜてでねめれあかさたなまら]ずに|[えけげせぜてでねめれいきしじちにみりっ]て(?=い(る|ま|く|け))|
    [一-龯]し?い(?=([一-龯]|[゠-ヿ]{2}))(?!出)|[一-龯]しく(?=([一-龯]|[゠-ヿ]{2}))|べきじゃ(?!あ)|かなり(?=([一-龯]|[゠-ヿ]{2}))|[えけげせぜてでねめれ]ば(?=([一-龯]|[゠-ヿ]{2}))|
    ゆっくり(?=([一-龯]|[゠-ヿ]{2}))|ちゃんと(?=([一-龯]|[゠-ヿ]{2}))|(なければ|なきゃ|ないと)(?=(なら|いけ))|[ぁ-ゖ゛-ゟ](?=(はず|べき)(だ|よ|$|。|…|！|？))|[ぁ-ゖ゛-ゟ](?=[゠-ヿ]{2})|て(?=ありがと)|
    なら(?=([一-龯]|[゠-ヿ]{2}))|なのは|[えけげせぜてでねめれ][るてた](?=([一-龯]|[゠-ヿ]{2}))|たく(?=な[いか])|[わかさたなまら]れ[るた](?=([一-龯]|[゠-ヿ]{2}))|いくつか|[一-龯]ても|して(?=([一-龯]|[゠-ヿ]{2}))|
    [一-龯]たる(?=([一-龯]|[゠-ヿ]{2}))|という(?=([一-龯]|[゠-ヿ]{2}))|を|な[くい](?=([一-龯]|[゠-ヿ]{2}))|[一-龯][ぁ-ゖ゛-ゟ]に(?=な(る|った|らな))|いた(?=([一-龯]|[゠-ヿ]{2}))|
    ないと(?=([一-龯]|[゠-ヿ]{2}))|て(?=ほし[いくか])|[一-龯]{2}(?=[゠-ヿ]{2})|な(?=([一-龯]|[゠-ヿ]{2}))|[゠-ヿ]{2}(?=[一-龯]{2})|(?P<doubler>[ぁ-ゖ゛-ゟ]{2})(?P=doubler)|くて(?=[一-龯])|
    しか(?=([一-龯]|[゠-ヿ]{2}))
)
""", re.VERBOSE)

# --- Functions ---

#def cleanup_zwsp_spacing(text):
#    """Remove extra ZWSPs that are within 1–2 chars of another, preserving text."""
#    if not isinstance(text, str):
#        return text
#    return re.sub(r'\u200B(.{1,2})\u200B', lambda m: m.group(1) + '\u200B', text)

def postprocess_ellipses(text):
    """Handle special rules for ellipses (… and ……):
       - No ZWSP if text starts with ellipsis
       - Add ZWSP after single '…' (not '……') when mid-sentence
       - Remove stray ZWSP snuck before ellipses
    """
    if not isinstance(text, str):
        return text

    # 1️⃣ Remove ZWSP immediately after starting ellipses
    text = re.sub(r'^(…{1,4})\u200B', r'\1', text)

    # 2️⃣ Add ZWSP after single ellipsis not followed by another ellipsis
    text = re.sub(r'(?<!…)(…)(?!…)(?=\S)', lambda m: m.group(1) + '\u200B', text)

    # 1️⃣ Remove stray ZWSP snuck before ellipses
    text = re.sub(r'([^\s…])\u200B(…|\.\.\.)', r'\1\2', text)

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
    #processed = cleanup_zwsp_spacing(processed)

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

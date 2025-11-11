import regex as re
from google.colab import files
import io
from openpyxl import load_workbook
import os

# ============================================================
#  MODE SELECTION
# ============================================================
print("Choose a mode:\n"
      "  1. Insert delimiters into an Excel file\n"
      "  2. Linebreak a Japanese text segment into balanced chunks\n")

mode = input("Enter 1 or 2 (default: 1): ").strip() or "1"

# --- Main regex ---
pattern = re.compile(r"""
(
    ([一-龯]{1,2}|[゠-ヿ]{2,12}|こと|ところ|[一-龯](?:[ぁ-ゖ゛-ゟ](?!で))+[一-龯]|[゠-ヿ]{2,12}[一-龯]|もの|入り|」|たち|ここ|そこ|[一-龯]ら|(?P<double>[ぁ-ゖ゛-ゟ]{2})(?P=double)|[えけげせぜてでねめれ]る|まま|[あこそ]いつ|あ[なん]た|さん|まみれ|おそらく|たっぷり|気持ち|すら|さすが|くず|あちこち|もと)
    (が(?!(して|った))|か(?!([はもらえけげせぜてでねめれいきぎしちにんをうくぐすつぬむるっ]|った|さ))|か[は]|は(?!ず)|も(?!の)|の(?![みにがはた為])|なく(?!て)|な(?![くのんらるい])|する(?!な)|から(?!して)|まで|に(?!([はも]|ついて|関して|すら))|
    に[はも]|へ[の]|へ(?![の])|で(?![はもすしきの])|で[はも]|じて(?!る)|や(?![かり])|と[のはか]|と(?!([のなはか]|[い言云]う))|して[はも]|して(?![はもる])|ならば|なら(?![ばで]))
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
    しか(?=([一-龯]|[゠-ヿ]{2}))|よりかは|て(?=しま[ういわ])
)
""", re.VERBOSE)

# ============================================================
#  MODE 1: Excel delimiter insertion
# ============================================================
if mode == "1":
    user_input = input("Enter a delimiter (press Enter for invisible ZWSP '\\u200B'): ").strip()
    INSERT_CHAR = user_input if user_input else '\u200B'
    preview_symbol = "[ZWSP]" if INSERT_CHAR == '\u200B' else INSERT_CHAR
    print(f"✅ Using delimiter: {repr(INSERT_CHAR)}")
    print(f"🔍 Preview: 日本語{preview_symbol}テキスト")

    uploaded = files.upload()
    filename = list(uploaded.keys())[0]
    wb = load_workbook(io.BytesIO(uploaded[filename]))
    target_headers = {"ja", "jp", "jap", "japanese"}

    def postprocess_ellipses(text):
        if not isinstance(text, str):
            return text
        text = re.sub(rf'^(…{{1,4}}){re.escape(INSERT_CHAR)}', r'\1', text)
        text = re.sub(r'(?<!…)(…)(?!…)(?=\S)', lambda m: m.group(1) + INSERT_CHAR, text)
        text = re.sub(rf'([^\s…]){re.escape(INSERT_CHAR)}(…|\.\.\.)', r'\1\2', text)
        return text

    def insert_delimiter(text):
        if not isinstance(text, str):
            return text
        def replacer(m):
            end = m.end()
            remainder = text[end:]
            next_char = remainder[:1]
            if re.match(r'[、。？！,．,.!?"”」』）)]', next_char) or re.match(r'^[、。？！…‥！？\s]*$', remainder):
                return m.group(0)
            return m.group(0) + INSERT_CHAR
        processed = pattern.sub(replacer, text)
        return postprocess_ellipses(processed)

    for ws in wb.worksheets:
        headers = {cell.value: cell.column for cell in ws[1] if cell.value}
        for header, col in headers.items():
            if str(header).strip().lower() in target_headers:
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col)
                    if isinstance(cell.value, str):
                        new_val = insert_delimiter(cell.value)
                        if new_val != cell.value:
                            cell.value = new_val

    name, ext = os.path.splitext(filename)
    output_filename = f"delimiters_added_{name}{ext}"
    wb.save(output_filename)
    files.download(output_filename)
    print(f"✅ Done! File saved as: {output_filename}")

# ============================================================
#  MODE 2: Smart text segment linebreaker
# ============================================================
elif mode == "2":
    text = input("Paste the Japanese text segment:\n").strip()
    lines = int(input("How many lines would you like to split it into? ").strip())

    # Find all potential breakpoints
    break_positions = [m.end() for m in pattern.finditer(text)]
    if not break_positions:
        print("⚠️ No suitable breakpoints found.")
    else:
        total_len = len(text)
        target_len = total_len / lines
        chosen_breaks = []
        last = 0

        for i in range(1, lines):
            target_pos = target_len * i
            best_break = min(break_positions, key=lambda x: abs(x - target_pos))
            if best_break > last:
                chosen_breaks.append(best_break)
                last = best_break

        # Remove duplicates and sort
        chosen_breaks = sorted(set(chosen_breaks))
        chunks = []
        prev = 0
        for bp in chosen_breaks:
            chunks.append(text[prev:bp])
            prev = bp
        chunks.append(text[prev:])

        print("\n✅ Suggested linebreaks:\n")
        for i, chunk in enumerate(chunks, 1):
            print(f"{i:02d}: {chunk}")

else:
    print("⚠️ Invalid mode. Exiting.")

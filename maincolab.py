import regex as re
from google.colab import files
import io
from openpyxl import load_workbook
import os

# ============================================================
# 🧩 MODE SELECTION
# ============================================================
print("Choose a mode:\n"
      "  1. Insert delimiters into an Excel file\n"
      "  2. Linebreak a Japanese text segment into balanced chunks\n")

mode = input("Enter 1 or 2 (default: 1): ").strip() or "1"

# --- Main regex ---
pattern = re.compile(r"""
(
    (\p{Han}{1,2}|\p{Katakana}{2,12}|こと|ところ|\p{Han}(?:\p{Hiragana}(?!で))+\p{Han}|\p{Katakana}{2,12}\p{Han}|もの|入り|」|たち|ここ|そこ|\p{Han}ら|(?P<double>\p{Hiragana}{2})(?P=double)|[えけげせぜてでねめれ]る|まま|[あこそ]いつ|あ[なん]た|さん|まみれ|おそらく|たっぷり|気持ち|すら|さすが|くず|あちこち|もと|さま)
    (が(?!(して|った))|か(?!([はもらえけげせぜてでねめれいきぎしちにんをうくぐすつぬむるっ]|った|さ))|か[は]|は(?!ず)|も(?!の)|の(?![みにがはた為])|なく(?!て)|な(?![くのんらるい])|する(?!な)|から(?!して)|まで|に(?!([はも]|ついて|関して|すら))|
    に[はも]|へ[の]|へ(?![の])|で(?![はもすしきの])|で[はも]|じて(?!る)|や(?![かり])|と[のはか]|と(?!([のなはか]|[い言云]う))|して[はも]|して(?![はもる])|ならば|なら(?![ばで]))
    |
    [、。？！・：；]
    |
    (――)|(……)|(\.\.\.)
    |
    \p{Han}すぎ[^たるだ]
    |
    について(?![はも])|について[はも]|に関して(?![はも])|に関して[はも]|[っいきぎしちにん][ただ]り|とにかく|でも|[くぐ]らい(?!は)|[くぐ]らいは|まるで|って(?![るたかも])|っても|
    すなわち|[うくぐすつぬふむる]の[にはもが]|を|んな[のに]|[って]たら|として|つまり|ちょっと|ちょうど|々な|々に(?![もは])|々に[もは]|たい(?=\p{Han})|けど|よう[なに](?=(\p{Han}{2}|\p{Katakana}{2}))|
    だと(?!は)|だとは|とは|[のただ]ほうが|ないほうが|[のただ]方が|ない方が|風に|[いきしちにひみり]たくて|[うくすつぬふむる]まて|[^一-龯]続く|ないと(?=いけ)|く(?=(\p{Han}|\p{Katakana}{2}))|
    ほとんど|らしくて(?!は)|らしく(?!て)|ため([にの](?![はも])|ならば|なら(?!ば))|ため[にの][はも]|為に(?![はも])|為に[はも]|わけ(では|じゃ(?!あ))|ほうが(?=(\p{Han}|\p{Katakana}{2}))|
    いきなり|すれば|(れば|ないと)(?=([い良善好]い|[よ良善好]か))|て(?=い?ました)|しっかり|して(?=あげ([るた]|(ます|まし)))|て(?=(ください|下さい|ちょうだい))|これまでに(?!は)|
    より(?=ずっと)|はじめて|て(?=くれ)|くなって(?!は)|され[るた](?![んの])|かった(?![んのりわっがぞぜ])|もなくて(?!は)|あらゆる|すべて(の|を|では|じゃ(?!あ))|すぐに[はも]|すぐに(?![はも])|
    もなく(?!て)|ながら|がてら|った(?![らんのりわっがぞぜ])|よりも|かも(?=[しれ])|とともに(?![はも])|と共に(?![はも])|もっとも|すべて(?!でのを)|ただの|まま(?=(\p{Han}|\p{Katakana}{2}))|
    どうして|どうやって|した(?=(\p{Han}{2}|こと|とこ))|のもとに|[うくすつぬふむるじの]よう[にな]|れて(?=(いき?ま|いる|いた|いな))|じゃ(?=な[いか])|では(?=な[いか])|またしても|
    どうなるか(?!は)|どうなるかは|しばらく|[えけげせぜてでねめれ]なく(?!て)|[えけげせぜてでねめれあかさたなまら]ずに|[えけげせぜてでねめれいきしじちにみりっ]て(?=い(る|ま|く|け))|
    \p{Han}し?い(?=(\p{Han}|\p{Katakana}{2}))(?!出)|\p{Han}しく(?=(\p{Han}|\p{Katakana}{2}))|べきじゃ(?!あ)|かなり(?=(\p{Han}|\p{Katakana}{2}))|[えけげせぜてでねめれ]ば(?=(\p{Han}|\p{Katakana}{2}))|
    ゆっくり(?=(\p{Han}|\p{Katakana}{2}))|ちゃんと(?=(\p{Han}|\p{Katakana}{2}))|(なければ|なきゃ|ないと)(?=(なら|いけ))|\p{Hiragana}(?=(はず|べき)(だ|よ|$|。|…|！|？))|\p{Hiragana}(?=\p{Katakana}{2})|て(?=ありがと)|
    なら(?=(\p{Han}|\p{Katakana}{2}))|なのは|[えけげせぜてでねめれ][るてた](?=(\p{Han}|\p{Katakana}{2}))|たく(?=な[いか])|[わかさたなまら]れ[るた](?=(\p{Han}|\p{Katakana}{2}))|いくつか|\p{Han}ても|して(?=(\p{Han}|\p{Katakana}{2}))|
    \p{Han}たる(?=(\p{Han}|\p{Katakana}{2}))|という(?=(\p{Han}|\p{Katakana}{2}))|を|な[くい](?=(\p{Han}|\p{Katakana}{2}))|\p{Han}\p{Hiragana}に(?=な(る|った|らな))|いた(?=(\p{Han}|\p{Katakana}{2}))|
    ないと(?=(\p{Han}|\p{Katakana}{2}))|て(?=ほし[いくか])|\p{Han}{2}(?=\p{Katakana}{2})|な(?=(\p{Han}|\p{Katakana}{2}))|\p{Katakana}{2}(?=\p{Han}{2})|(?P<doubler>\p{Hiragana}{2})(?P=doubler)|くて(?=\p{Han})|
    しか(?=(\p{Han}|\p{Katakana}{2}))|よりかは|て(?=しま[ういわ])|とっ?ても
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
            # ensure break moves forward to avoid infinite loops
            if best_break > last:
                chosen_breaks.append(best_break)
                last = best_break

        chosen_breaks = sorted(set(chosen_breaks))
        chunks = []
        prev = 0
        for bp in chosen_breaks:
            chunks.append(text[prev:bp])
            prev = bp
        chunks.append(text[prev:])

        # --- Polishing pass: punctuation + short-token fixes ---
        def polish_lines(chunks):
            """Avoid lines starting/ending with dangling punctuation or short 'conjunct + punctuation' heads."""
            adjusted = chunks[:]  # work on a copy
            punct = "、。？！：；…‥" + "\.\.\."
            # 1) Move trailing punctuation (within last 1-3 chars) to next line
            for i in range(len(adjusted) - 1):
                for n in range(1, 4):
                    if len(adjusted[i]) >= n and adjusted[i][-n] in punct:
                        # move those n chars to start of next line
                        adjusted[i+1] = adjusted[i][-n:] + adjusted[i+1]
                        adjusted[i] = adjusted[i][:-n]
                        break

            # 2) Move leading punctuation to previous line
            for i in range(1, len(adjusted)):
                for n in range(1, 4):
                    if len(adjusted[i]) >= n and adjusted[i][0] in punct:
                        adjusted[i-1] += adjusted[i][:n]
                        adjusted[i] = adjusted[i][n:]
                        break

            # 3) Move punctuation within first 1-3 chars to previous line
            for i in range(1, len(adjusted)):
                m = re.match(r'^([\p{Hiragana}\p{Katakana}\p{Han}]{1,3})([、。？！…])', adjusted[i])
                if m:
                    tok = m.group(1) + m.group(2)
                    # move token to previous line, avoid creating empty previous line
                    adjusted[i-1] += tok
                    adjusted[i] = adjusted[i][len(tok):]

            # final pass: trim accidental empty lines (but keep at least one char if possible)
            final = []
            for part in adjusted:
                if part == "" and final:
                    # if empty and there's a previous, merge with previous to avoid empties
                    final[-1] += ""
                else:
                    final.append(part)
            return final

        chunks = polish_lines(chunks)

        print("\n✅ Suggested linebreaks:\n")
        for i, chunk in enumerate(chunks, 1):
            print(f"{i:02d}: {chunk}")

else:
    print("⚠️ Invalid mode. Exiting.")

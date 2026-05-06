import argparse
import io
import os
import re
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape as xml_escape
from typing import Iterable

import pandas as pd
import pymorphy3
import pythoncom
import win32com.client
from openpyxl import load_workbook
from docx import Document
from docx.document import Document as _Document
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from xml.etree import ElementTree as ET

DEFAULT_SHEET_INDEX = 1  # second sheet
DEFAULT_NAME_COL_INDEX = 1  # second column (B)
DEFAULT_COMPANY_COL_INDEX = 0  # first column (A)
DEFAULT_ROLE_COL_INDEX = 2  # third column (C)
DEFAULT_TARGET_TEXT = "Когогину С.А."
DEFAULT_COMPANY_TARGET_TEXT = "ПАО «Камаз»"
DEFAULT_ROLE_TARGET_TEXT = "Генеральному директору"
DEFAULT_EMAIL_TARGET_TEXT = "info@truckinstock.com"
DEFAULT_SALUTATION_ANCHOR = "Уважаемый "
DEFAULT_TARGET_CASE = "datv"  # дательный падеж (пример: Ширманов -> Ширманову)
EMAIL_REGEX = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
MORPH = pymorphy3.MorphAnalyzer()


def inflect_to_genitive(word: str) -> str:
    """
    Convert surname to recipient form (dative) when possible.
    Keeps punctuation around the word unchanged.
    """
    match = re.match(r"^([^A-Za-zА-Яа-яЁё]*)([A-Za-zА-Яа-яЁё-]+)([^A-Za-zА-Яа-яЁё]*)$", word)
    if not match:
        return word

    prefix, core, suffix = match.groups()
    parses = MORPH.parse(core)
    if not parses:
        return word

    # Prefer surname-like singular parses for correct family-name inflection.
    preferred = None
    core_lower = core.lower()
    prefer_female_surname = core_lower.endswith(
        ("ова", "ева", "ёва", "ина", "ына", "ая", "яя")
    )

    if prefer_female_surname:
        for p in parses:
            tag_str = str(p.tag)
            if "Surn" in tag_str and "sing" in tag_str and "femn" in tag_str:
                preferred = p
                break
    for p in parses:
        if preferred is not None:
            break
        tag_str = str(p.tag)
        if "Surn" in tag_str and "sing" in tag_str:
            preferred = p
            break
    if preferred is None:
        for p in parses:
            tag_str = str(p.tag)
            if "sing" in tag_str and ("NOUN" in tag_str or "NPRO" in tag_str):
                preferred = p
                break
    if preferred is None:
        preferred = parses[0]

    inflected = preferred.inflect({"datv"})
    inflected_core = inflected.word if inflected else core

    # Fallback heuristic for surnames if parser returns plural-looking output.
    if inflected_core.endswith("ам") and (
        core_lower.endswith("ов")
        or core_lower.endswith("ев")
        or core_lower.endswith("ин")
        or core_lower.endswith("ын")
    ):
        inflected_core = f"{core_lower}у"

    # Fallback for common surname endings when inflection is unchanged.
    if inflected_core.lower() == core_lower:
        if core_lower.endswith("ов") or core_lower.endswith("ев") or core_lower.endswith("ёв"):
            inflected_core = f"{core_lower}у"
        elif core_lower.endswith("ин") or core_lower.endswith("ын"):
            inflected_core = f"{core_lower}у"
        elif core_lower.endswith("ова"):
            inflected_core = f"{core_lower[:-1]}ой"
        elif core_lower.endswith("ева"):
            inflected_core = f"{core_lower[:-1]}ой"
        elif core_lower.endswith("ёва"):
            inflected_core = f"{core_lower[:-1]}ой"
        elif core_lower.endswith("ина") or core_lower.endswith("ына"):
            inflected_core = f"{core_lower[:-1]}ой"
        elif core_lower.endswith("ая"):
            inflected_core = f"{core_lower[:-2]}ой"

    # pymorphy обычно возвращает в нижнем регистре; сохраняем стиль исходного слова.
    if core.isupper():
        inflected_core = inflected_core.upper()
    elif core[:1].isupper():
        inflected_core = inflected_core[:1].upper() + inflected_core[1:]

    return f"{prefix}{inflected_core}{suffix}"


def read_excel_value(
    excel_path: str,
    sheet_index: int,
    col_index: int,
    row_number: int,
) -> str:
    """
    Read a single cell from Excel using:
    - sheet_index: zero-based sheet index (1 => second sheet)
    - col_index: zero-based column index (1 => second column, 'B')
    - row_number: Excel row number (2 => B2)

    Assumes headers are on the first Excel row.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    # Load the target sheet
    xls = pd.ExcelFile(excel_path)
    if sheet_index < 0 or sheet_index >= len(xls.sheet_names):
        raise IndexError(
            f"sheet_index {sheet_index} out of range for workbook "
            f"(sheets: {xls.sheet_names})"
        )

    sheet_name = xls.sheet_names[sheet_index]
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Excel row 1 is headers; data rows start at Excel row 2 => df index 0
    df_row = row_number - 2
    if df_row < 0 or df_row >= len(df):
        raise IndexError(
            f"Excel row_number {row_number} maps to df index {df_row}, "
            f"which is outside the data range 0..{len(df) - 1}"
        )

    if col_index < 0 or col_index >= df.shape[1]:
        raise IndexError(
            f"col_index {col_index} out of range for sheet "
            f"with {df.shape[1]} columns"
        )

    value = df.iat[df_row, col_index]
    if pd.isna(value):
        raise ValueError(
            f"Cell at sheet_index={sheet_index}, col_index={col_index}, "
            f"row_number={row_number} is empty (NaN)"
        )

    return str(value)


def load_sheet_dataframe(excel_path: str, sheet_index: int) -> pd.DataFrame:
    """
    Load a sheet by zero-based index from Excel.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    xls = pd.ExcelFile(excel_path)
    if sheet_index < 0 or sheet_index >= len(xls.sheet_names):
        raise IndexError(
            f"sheet_index {sheet_index} out of range for workbook "
            f"(sheets: {xls.sheet_names})"
        )
    sheet_name = xls.sheet_names[sheet_index]
    return pd.read_excel(excel_path, sheet_name=sheet_name)


def load_sheet_for_styles(excel_path: str, sheet_index: int):
    """
    Load Excel worksheet with style information.
    """
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.worksheets[sheet_index]
    return wb, ws


def is_cell_yellow(cell) -> bool:
    """
    Detect yellow fill in an openpyxl cell.
    """
    fill = cell.fill
    if fill is None:
        return False
    if fill.fill_type in (None, "none"):
        return False
    color = fill.fgColor
    if color is None:
        return False

    # RGB/ARGB yellow variants
    if color.type == "rgb" and color.rgb:
        rgb = color.rgb.upper()
        return rgb in ("FFFF00", "FFFFFF00")

    # Indexed palette yellow is commonly index 6
    if color.type == "indexed" and color.indexed is not None:
        return int(color.indexed) == 6

    return False


def is_row_yellow(ws, excel_row_number: int) -> bool:
    """
    True if any cell in the Excel row has yellow fill.
    """
    for cell in ws[excel_row_number]:
        if is_cell_yellow(cell):
            return True
    return False


def extract_first_email(text: str) -> str:
    """
    Extract first email found anywhere in text.
    """
    match = EMAIL_REGEX.search(str(text))
    if not match:
        raise ValueError("No email found in text")
    return match.group(0)


def format_replacement_name(raw_value: str) -> str:
    """
    Transform source text:
    - take at most first 3 words.
    - if 3 words: first two -> initials ("X.Y."), third stays full.
    - if 2 words: first -> initial ("X."), second stays full.
    - if 1 word: keep it full.
    - final output word is converted to genitive case.
    """
    words = [w for w in str(raw_value).split() if w]
    words = words[:3]
    if not words:
        raise ValueError("Source cell has no words to format")

    result_parts = []
    if len(words) >= 3:
        surname_gen = inflect_to_genitive(words[2])
        return f"{words[0][0]}.{words[1][0]}.{surname_gen}"
    elif len(words) == 2:
        surname_gen = inflect_to_genitive(words[1])
        return f"{words[0][0]}.{surname_gen}"
    else:
        result_parts = [words[0]]

    result_parts[-1] = inflect_to_genitive(result_parts[-1])

    return " ".join(result_parts)


def normalize_fio_order(raw_value: str) -> str:
    """
    Normalize 3-word full name to order: FirstName Patronymic LastName.
    If confidence is low, keep original order.
    """
    words = [w for w in str(raw_value).split() if w]
    if len(words) != 3:
        return str(raw_value).strip()

    def is_patronymic(word: str) -> bool:
        wl = word.lower().strip(".,;:!?()[]{}\"'")
        return (
            wl.endswith("вич")
            or wl.endswith("ич")
            or wl.endswith("вна")
            or wl.endswith("ична")
            or wl.endswith("инична")
        )

    def name_score(word: str) -> int:
        score = 0
        parses = MORPH.parse(word)
        for p in parses[:5]:
            tag = str(p.tag)
            if "Name" in tag:
                score += 3
            if "Surn" in tag:
                score -= 2
        return score

    def surname_score(word: str) -> int:
        score = 0
        parses = MORPH.parse(word)
        for p in parses[:5]:
            tag = str(p.tag)
            if "Surn" in tag:
                score += 3
            if "Name" in tag:
                score -= 1
        wl = word.lower().strip(".,;:!?()[]{}\"'")
        if wl.endswith(("ов", "ев", "ёв", "ин", "ын", "ская", "ская", "ская", "енко", "ук", "юк")):
            score += 1
        return score

    # 1) Locate patronymic confidently.
    patronymic_idx = None
    for i, w in enumerate(words):
        if is_patronymic(w):
            patronymic_idx = i
            break
    if patronymic_idx is None:
        # no strong signal -> keep original
        return " ".join(words)

    remaining = [i for i in range(3) if i != patronymic_idx]
    i1, i2 = remaining[0], remaining[1]

    # 2) Decide which remaining token is first name vs surname.
    first_candidate = i1
    last_candidate = i2
    n1, n2 = name_score(words[i1]), name_score(words[i2])
    s1, s2 = surname_score(words[i1]), surname_score(words[i2])

    # Prefer strongest evidence.
    if (n2 - s2) > (n1 - s1):
        first_candidate, last_candidate = i2, i1

    ordered = [words[first_candidate], words[patronymic_idx], words[last_candidate]]
    return " ".join(ordered)


def _preserve_word_case(source: str, transformed: str) -> str:
    if source.isupper():
        return transformed.upper()
    if source[:1].isupper():
        return transformed[:1].upper() + transformed[1:]
    return transformed


def _inflect_word_to_case(word: str, target_case: str, prefer_name: bool = False) -> str:
    """
    Inflect a single Russian word to the target case, preserving source casing.
    """
    parses = MORPH.parse(word)
    if not parses:
        return word

    preferred = None
    if prefer_name:
        for p in parses:
            tag_str = str(p.tag)
            if "Name" in tag_str and "sing" in tag_str:
                preferred = p
                break
        if preferred is None:
            for p in parses:
                tag_str = str(p.tag)
                if "Patr" in tag_str and "sing" in tag_str:
                    preferred = p
                    break

    if preferred is None:
        for p in parses:
            tag_str = str(p.tag)
            if "sing" in tag_str and ("NOUN" in tag_str or "ADJF" in tag_str or "PRTF" in tag_str):
                preferred = p
                break
    if preferred is None:
        preferred = parses[0]

    inflected = preferred.inflect({target_case})
    if inflected is None:
        return word
    return _preserve_word_case(word, inflected.word)


def _inflect_word_to_dative(word: str) -> str:
    """
    Inflect a single Russian word to dative when possible.
    """
    return _inflect_word_to_case(word, DEFAULT_TARGET_CASE)


def format_salutation_name(raw_value: str) -> str:
    """
    Build salutation name fragment:
    - if 2 words: take only first word (name), no case inflection;
    - if 3+ words: take first + second words, no case inflection.
    """
    words = [w for w in str(raw_value).split() if w]
    if not words:
        raise ValueError("Name cell is empty for salutation")

    if len(words) == 1:
        selected = [words[0]]
    elif len(words) == 2:
        selected = [words[0]]
    else:
        selected = [words[0], words[1]]

    return " ".join(selected)


def detect_gender_from_name(raw_value: str) -> str | None:
    """
    Detect person gender from full name text.
    Returns:
    - "femn" for female
    - "masc" for male
    - None if unknown
    """
    words = [w.strip(".,;:!?()[]{}\"'") for w in str(raw_value).split() if w]
    if not words:
        return None

    # 1) Patronymic heuristic is the most reliable in Russian names.
    for w in words:
        wl = w.lower()
        if wl.endswith("вна") or wl.endswith("ична") or wl.endswith("инична"):
            return "femn"
        if wl.endswith("ич"):
            return "masc"

    # 2) Morphology fallback: try to detect from likely name words.
    for w in words[:3]:
        parses = MORPH.parse(w)
        for p in parses:
            tag_str = str(p.tag)
            if "Name" in tag_str:
                if "femn" in tag_str:
                    return "femn"
                if "masc" in tag_str:
                    return "masc"

    return None


def get_salutation_prefix(raw_value: str) -> str:
    """
    Return gender-aware salutation prefix.
    """
    gender = detect_gender_from_name(raw_value)
    if gender == "femn":
        return "Уважаемая"
    return "Уважаемый"


def format_role_to_dative(raw_value: str) -> str:
    """
    Convert role/title text to dative in a conservative way:
    - work by slash-separated segments;
    - skip segments that start with "И.о.";
    - in each segment, inflect only the leading nominative phrase:
      stop at first preposition or first non-nominative/accusative word.
    This matches cases like:
    - "Начальник отдела ..." -> "Начальнику отдела ..."
    - "первый проректор по ..." -> "первому проректору по ..."
    """
    text = str(raw_value).strip()
    if not text:
        raise ValueError("Role cell is empty")

    segments = re.split(r"(\s*/\s*)", text)
    result_segments = []

    for seg in segments:
        if "/" in seg:
            result_segments.append(seg)
            continue

        if re.match(r"^\s*и\.о\.", seg, flags=re.IGNORECASE):
            result_segments.append(seg)
            continue

        tokens = re.split(r"(\s+)", seg)
        out_tokens = []
        started = False
        done = False

        for tok in tokens:
            if not tok or tok.isspace():
                out_tokens.append(tok)
                continue

            m = re.match(r"^([^A-Za-zА-Яа-яЁё-]*)([A-Za-zА-Яа-яЁё-]+)([^A-Za-zА-Яа-яЁё-]*)$", tok)
            if not m:
                out_tokens.append(tok)
                continue

            prefix, core, suffix = m.groups()
            parses = MORPH.parse(core)
            if not parses:
                out_tokens.append(tok)
                continue

            p = parses[0]
            tag = p.tag
            pos = tag.POS

            if done:
                out_tokens.append(tok)
                continue

            # Stop inflection on prepositions.
            if pos == "PREP":
                done = True
                out_tokens.append(tok)
                continue

            # Start only on adjective/noun-like words.
            if not started:
                if pos not in ("NOUN", "ADJF", "PRTF", "NPRO"):
                    out_tokens.append(tok)
                    continue
                started = True

            case = tag.case
            # Only keep inflecting while words look like nominative/accusative.
            if case not in ("nomn", "accs", None):
                done = True
                out_tokens.append(tok)
                continue

            out_core = _inflect_word_to_dative(core)
            out_tokens.append(f"{prefix}{out_core}{suffix}")

        result_segments.append("".join(out_tokens))

    return "".join(result_segments)


def _iter_paragraphs_from_table(table: Table) -> Iterable[Paragraph]:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                yield paragraph
            for nested_table in cell.tables:
                yield from _iter_paragraphs_from_table(nested_table)


def _iter_all_paragraphs(doc: _Document) -> Iterable[Paragraph]:
    # Body
    for paragraph in doc.paragraphs:
        yield paragraph
    for table in doc.tables:
        yield from _iter_paragraphs_from_table(table)

    # Headers/footers per section
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            yield paragraph
        for table in section.header.tables:
            yield from _iter_paragraphs_from_table(table)
        for paragraph in section.footer.paragraphs:
            yield paragraph
        for table in section.footer.tables:
            yield from _iter_paragraphs_from_table(table)


def _replace_in_runs(runs, start: int, end: int, replacement: str) -> None:
    """
    Replace [start:end) in concatenated run text for a given run sequence.
    """
    run_texts = [r.text or "" for r in runs]
    if not run_texts:
        return

    starts = [0]
    for txt in run_texts[:-1]:
        starts.append(starts[-1] + len(txt))
    ends = [s + len(t) for s, t in zip(starts, run_texts)]

    si = next(i for i, (a, b) in enumerate(zip(starts, ends)) if a <= start < b)
    ei = next(i for i, (a, b) in enumerate(zip(starts, ends)) if a <= end - 1 < b)

    if si == ei:
        local_s = start - starts[si]
        local_e = end - starts[si]
        runs[si].text = run_texts[si][:local_s] + replacement + run_texts[si][local_e:]
        return

    local_s = start - starts[si]
    runs[si].text = run_texts[si][:local_s] + replacement

    local_e_last = end - starts[ei]
    runs[ei].text = run_texts[ei][local_e_last:]

    for k in range(si + 1, ei):
        runs[k].text = ""


def _register_original_namespaces(xml_bytes: bytes) -> None:
    seen = set()
    for _, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
        key = (prefix or "", uri)
        if key in seen:
            continue
        seen.add(key)
        ET.register_namespace(prefix or "", uri)


def _rewrite_docx_with_document_xml(
    source_docx: str,
    output_path: str,
    new_document_xml: bytes,
) -> None:
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with zipfile.ZipFile(source_docx, "r") as zin, zipfile.ZipFile(
        output_path, "w"
    ) as zout:
        for item in zin.infolist():
            data = (
                new_document_xml
                if item.filename == "word/document.xml"
                else zin.read(item.filename)
            )
            zout.writestr(item, data)


def _replace_text_raw_in_docx_xml_parts(
    docx_path: str,
    target: str,
    replacement: str,
    output_path: str,
) -> int:
    """
    Replace text directly in XML payloads without parsing/reserializing XML.
    This is the safest path for Word compatibility when replacement is exact text.
    """
    replacements = 0
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(
        output_path, "w"
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                try:
                    txt = data.decode("utf-8")
                except UnicodeDecodeError:
                    txt = None
                if txt is not None:
                    cnt = txt.count(target)
                    if cnt:
                        safe_replacement = xml_escape(replacement)
                        txt = txt.replace(target, safe_replacement)
                        data = txt.encode("utf-8")
                        replacements += cnt
            zout.writestr(item, data)

    return replacements


def _replace_text_in_document_xml_by_paragraph(
    docx_path: str,
    target: str,
    replacement: str,
    output_path: str,
) -> int:
    with zipfile.ZipFile(docx_path, "r") as z:
        xml_bytes = z.read("word/document.xml")

    _register_original_namespaces(xml_bytes)
    root = ET.fromstring(xml_bytes)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    replacements = 0

    for p in root.findall(".//w:p", ns):
        t_nodes = p.findall(".//w:t", ns)
        if not t_nodes:
            continue
        texts = [(t.text or "") for t in t_nodes]
        para_text = "".join(texts)
        if not para_text:
            continue

        start_search = 0
        spans = []
        while True:
            idx = para_text.find(target, start_search)
            if idx == -1:
                break
            spans.append((idx, idx + len(target)))
            start_search = idx + len(target)
        if not spans:
            continue

        starts = [0]
        for txt in texts[:-1]:
            starts.append(starts[-1] + len(txt))
        ends = [s + len(t) for s, t in zip(starts, texts)]

        for start, end in reversed(spans):
            si = next(
                i for i, (a, b) in enumerate(zip(starts, ends)) if a <= start < b
            )
            ei = next(
                i for i, (a, b) in enumerate(zip(starts, ends)) if a <= end - 1 < b
            )
            if si == ei:
                local_s = start - starts[si]
                local_e = end - starts[si]
                t_nodes[si].text = texts[si][:local_s] + replacement + texts[si][local_e:]
            else:
                local_s = start - starts[si]
                t_nodes[si].text = texts[si][:local_s] + replacement
                local_e_last = end - starts[ei]
                t_nodes[ei].text = texts[ei][local_e_last:]
                for k in range(si + 1, ei):
                    t_nodes[k].text = ""
            replacements += 1

            # Refresh cached texts after each mutation for subsequent spans
            texts = [(t.text or "") for t in t_nodes]
            starts = [0]
            for txt in texts[:-1]:
                starts.append(starts[-1] + len(txt))
            ends = [s + len(t) for s, t in zip(starts, texts)]

    xml_new = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    _rewrite_docx_with_document_xml(docx_path, output_path, xml_new)
    return replacements


def replace_text_in_docx(
    docx_path: str,
    target: str,
    replacement: str,
    output_path: str,
) -> None:
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"Word file not found: {docx_path}")

    # Primary mode: preserve original XML structure exactly.
    replacements = _replace_text_raw_in_docx_xml_parts(
        docx_path=docx_path,
        target=target,
        replacement=replacement,
        output_path=output_path,
    )
    if replacements > 0:
        return

    # Fallback for split text cases where target isn't contiguous in XML.
    replacements = _replace_text_in_document_xml_by_paragraph(
        docx_path=docx_path,
        target=target,
        replacement=replacement,
        output_path=output_path,
    )
    if replacements == 0:
        raise ValueError(f"Target text not found: {target}")


def replace_text_with_word_com(
    docx_path: str,
    output_path: str,
    target: str,
    replacement: str,
) -> None:
    """
    Use Microsoft Word automation for replacement and save.
    This is the most compatibility-safe path for DOCX files.
    """
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    def replace_in_range(text_range) -> None:
        finder = text_range.Find
        finder.ClearFormatting()
        finder.Replacement.ClearFormatting()
        finder.Execute(
            FindText=target,
            MatchCase=False,
            MatchWholeWord=False,
            MatchWildcards=False,
            MatchSoundsLike=False,
            MatchAllWordForms=False,
            Forward=True,
            Wrap=1,  # wdFindContinue
            Format=False,
            ReplaceWith=replacement,
            Replace=2,  # wdReplaceAll
        )

    try:
        doc = word.Documents.Open(
            FileName=docx_path,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            OpenAndRepair=True,
        )

        # Main story
        replace_in_range(doc.Content)

        # Other stories: headers/footers/text frames/etc
        story = doc.StoryRanges
        if story is not None:
            current = story.Item(1)
            while current is not None:
                replace_in_range(current)
                next_story = current.NextStoryRange
                current = next_story

        # Shapes and inline shapes text frames
        for shp in doc.Shapes:
            if shp.TextFrame.HasText:
                replace_in_range(shp.TextFrame.TextRange)
        for shp in doc.InlineShapes:
            try:
                if shp.TextFrame.HasText:
                    replace_in_range(shp.TextFrame.TextRange)
            except Exception:
                pass

        # Save as DOCX (wdFormatXMLDocument = 12)
        doc.SaveAs2(output_path, FileFormat=12)
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()
        pythoncom.CoUninitialize()


def replace_texts_with_word_com(
    docx_path: str,
    output_path: str,
    replacements: list[tuple[str, str]],
) -> None:
    """
    Use Microsoft Word automation to apply multiple replacements and save once.
    """
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None

    def replace_salutation_by_place(text_range, replacement: str) -> None:
        """
        Find first salutation anchor and replace up to '!' in the same paragraph.
        """
        finder = text_range.Find
        finder.ClearFormatting()
        finder.Replacement.ClearFormatting()
        found = finder.Execute(
            FindText=DEFAULT_SALUTATION_ANCHOR,
            MatchCase=False,
            MatchWholeWord=False,
            MatchWildcards=False,
            MatchSoundsLike=False,
            MatchAllWordForms=False,
            Forward=True,
            Wrap=1,  # wdFindContinue
            Format=False,
        )
        if not found:
            return

        start_pos = text_range.Start
        para_end = text_range.Paragraphs(1).Range.End
        para_range = doc.Range(start_pos, para_end)
        para_text = para_range.Text

        excl_idx = para_text.find("!")
        if excl_idx == -1:
            return

        end_pos = start_pos + excl_idx + 1
        replace_range = doc.Range(start_pos, end_pos)
        replace_range.Text = replacement

    def replace_in_range(text_range) -> None:
        for item in replacements:
            if isinstance(item, tuple):
                target, replacement = item
                match_wildcards = False
                replace_mode = 2  # wdReplaceAll
            else:
                mode = item.get("mode")
                if mode == "salutation_place":
                    replace_salutation_by_place(text_range, item["replacement"])
                    continue
                target = item["target"]
                replacement = item["replacement"]
                match_wildcards = bool(item.get("wildcards", False))
                replace_mode = int(item.get("replace_mode", 2))

            finder = text_range.Find
            finder.ClearFormatting()
            finder.Replacement.ClearFormatting()
            finder.Execute(
                FindText=target,
                MatchCase=False,
                MatchWholeWord=False,
                MatchWildcards=match_wildcards,
                MatchSoundsLike=False,
                MatchAllWordForms=False,
                Forward=True,
                Wrap=1,  # wdFindContinue
                Format=False,
                ReplaceWith=replacement,
                Replace=replace_mode,
            )

    try:
        doc = word.Documents.Open(
            FileName=docx_path,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            OpenAndRepair=True,
        )
        replace_in_range(doc.Content)

        story = doc.StoryRanges
        if story is not None:
            current = story.Item(1)
            while current is not None:
                replace_in_range(current)
                current = current.NextStoryRange

        for shp in doc.Shapes:
            if shp.TextFrame.HasText:
                replace_in_range(shp.TextFrame.TextRange)
        for shp in doc.InlineShapes:
            try:
                if shp.TextFrame.HasText:
                    replace_in_range(shp.TextFrame.TextRange)
            except Exception:
                pass

        doc.SaveAs2(output_path, FileFormat=12)
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()
        pythoncom.CoUninitialize()


def replace_span_in_docx(
    docx_path: str,
    start_index: int,
    length: int,
    replacement: str,
    output_path: str,
) -> None:
    """Replace text by absolute position in concatenated paragraph text."""
    if start_index < 0:
        raise ValueError("start_index must be >= 0")
    if length <= 0:
        raise ValueError("length must be > 0")

    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"Word file not found: {docx_path}")

    doc = Document(docx_path)
    paragraphs = [p for p in _iter_all_paragraphs(doc) if p.runs]
    para_texts = ["".join(r.text or "" for r in p.runs) for p in paragraphs]
    combined = "".join(para_texts)

    end_index = start_index + length
    if end_index > len(combined):
        raise IndexError(
            f"Requested span [{start_index}, {end_index}) exceeds "
            f"combined text length {len(combined)}"
        )

    cursor = 0
    for paragraph, text in zip(paragraphs, para_texts):
        para_start = cursor
        para_end = cursor + len(text)
        cursor = para_end

        overlap_start = max(start_index, para_start)
        overlap_end = min(end_index, para_end)
        if overlap_start >= overlap_end:
            continue

        local_start = overlap_start - para_start
        local_end = overlap_end - para_start
        para_replacement = replacement if overlap_start == start_index else ""
        _replace_in_runs(paragraph.runs, local_start, local_end, para_replacement)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)


def build_default_paths() -> tuple[str, str]:
    """
    Build default paths for Excel and Word files in the current project.
    """
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

    # Default Excel: first non-temporary .xlsx in the base directory
    excel_candidates = [
        os.path.join(base_dir, f)
        for f in os.listdir(base_dir)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    ]
    if not excel_candidates:
        raise FileNotFoundError(
            f"No .xlsx files found in {base_dir}. "
            "Pass --excel explicitly."
        )
    excel_path = sorted(excel_candidates, key=os.path.getsize)[-1]

    # Default Word: largest non-temporary .docx in the base directory
    docx_candidates = [
        os.path.join(base_dir, f)
        for f in os.listdir(base_dir)
        if f.lower().endswith(".docx")
        and ".updated." not in f
        and ".pass2." not in f
        and not f.startswith("~$")
    ]
    if not docx_candidates:
        raise FileNotFoundError(
            f"No .docx files found in {base_dir}. "
            "Pass --word explicitly."
        )
    word_path = sorted(docx_candidates, key=os.path.getsize)[-1]

    return excel_path, word_path


def build_replacements_for_row(df: pd.DataFrame, row_idx: int) -> list:
    """
    Build all replacement values for one dataframe row.
    """
    replacement_raw = normalize_fio_order(str(df.iat[row_idx, DEFAULT_NAME_COL_INDEX]))
    replacement = format_replacement_name(replacement_raw)
    company_replacement = str(df.iat[row_idx, DEFAULT_COMPANY_COL_INDEX])
    role_raw = str(df.iat[row_idx, DEFAULT_ROLE_COL_INDEX])
    role_replacement = format_role_to_dative(role_raw)
    last_col_idx = df.shape[1] - 1
    email_replacement = extract_first_email(str(df.iat[row_idx, last_col_idx]))
    salutation_name = format_salutation_name(replacement_raw)
    salutation_prefix = get_salutation_prefix(replacement_raw)
    salutation_replacement = f"{salutation_prefix} {salutation_name}!"

    return [
        (DEFAULT_TARGET_TEXT, replacement),
        (DEFAULT_COMPANY_TARGET_TEXT, company_replacement),
        (DEFAULT_ROLE_TARGET_TEXT, role_replacement),
        (DEFAULT_EMAIL_TARGET_TEXT, email_replacement),
        {
            "mode": "salutation_place",
            "replacement": salutation_replacement,
        },
    ]


def validate_row_for_batch(df: pd.DataFrame, row_idx: int, ws=None) -> list[str]:
    """
    Validate row for batch processing.
    Rules:
    1) Column 2 must contain exactly 3 words.
    2) Last column must contain an email.
    """
    reasons = []
    name_value = str(df.iat[row_idx, DEFAULT_NAME_COL_INDEX]).strip()
    name_words = [w for w in name_value.split() if w]
    if len(name_words) != 3:
        reasons.append("Column 2 must contain exactly 3 words")

    last_col_idx = df.shape[1] - 1
    email_value = str(df.iat[row_idx, last_col_idx]).strip()
    if not EMAIL_REGEX.search(email_value):
        reasons.append("Last column must contain a valid email")

    if ws is not None:
        excel_row_number = row_idx + 2
        if is_row_yellow(ws, excel_row_number):
            reasons.append("Row has yellow fill")

    return reasons


def build_output_folder(base_dir: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = os.path.join(base_dir, f"output_{timestamp}")
    os.makedirs(folder, exist_ok=True)
    return folder


def process_single_row(
    df: pd.DataFrame,
    row_idx: int,
    default_word: str,
    output_folder: str,
) -> str:
    """
    Process one row and write one output Word file.
    """
    _, word_name = os.path.split(default_word)
    base, ext = os.path.splitext(word_name)
    excel_row_number = row_idx + 2
    output_name = f"{base}.updated.row{excel_row_number}{ext}"
    output_path = os.path.join(output_folder, output_name)

    replacements = build_replacements_for_row(df, row_idx)
    replace_texts_with_word_com(
        docx_path=default_word,
        output_path=output_path,
        replacements=replacements,
    )
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Update the default Word text location using a value from the "
            "default Excel source. Only row number is required."
        )
    )
    mode_group = parser.add_mutually_exclusive_group(required=True)
    mode_group.add_argument(
        "--rowNumber",
        type=int,
        help="Excel row number (2 => value from B2 on the second sheet).",
    )
    mode_group.add_argument(
        "--allRows",
        action="store_true",
        help="Process all data rows with validation rules.",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_args()
    default_excel, default_word = build_default_paths()
    word_dir, _ = os.path.split(default_word)
    output_folder = build_output_folder(word_dir)
    df = load_sheet_dataframe(default_excel, DEFAULT_SHEET_INDEX)

    if args.allRows:
        wb, ws = load_sheet_for_styles(default_excel, DEFAULT_SHEET_INDEX)
        processed_files = []
        skipped_rows = []
        try:
            for row_idx in range(len(df)):
                reasons = validate_row_for_batch(df, row_idx, ws=ws)
                if reasons:
                    row_data = df.iloc[row_idx].to_dict()
                    row_data["ExcelRow"] = row_idx + 2
                    row_data["Reason"] = "; ".join(reasons)
                    skipped_rows.append(row_data)
                    continue

                out_file = process_single_row(
                    df=df,
                    row_idx=row_idx,
                    default_word=default_word,
                    output_folder=output_folder,
                )
                processed_files.append(out_file)
        finally:
            wb.close()

        skipped_report_path = os.path.join(output_folder, "skipped_rows.xlsx")
        skipped_df = pd.DataFrame(skipped_rows)
        skipped_df.to_excel(skipped_report_path, index=False)

        print("Engine used: word-com")
        print("Mode: allRows")
        print("Output folder:", output_folder)
        print("Processed files:", len(processed_files))
        print("Skipped rows:", len(skipped_rows))
        print("Skipped report:", skipped_report_path)
        return

    # Single-row mode
    row_idx = args.rowNumber - 2
    if row_idx < 0 or row_idx >= len(df):
        raise IndexError(
            f"rowNumber {args.rowNumber} is out of range for available data rows"
        )
    output_path = process_single_row(
        df=df,
        row_idx=row_idx,
        default_word=default_word,
        output_folder=output_folder,
    )
    print("Engine used: word-com")
    print("Mode: rowNumber")
    print("Output folder:", output_folder)
    print("Updated Word file written to:", output_path)


if __name__ == "__main__":
    main()


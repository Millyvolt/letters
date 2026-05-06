import argparse
import io
import os
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape as xml_escape
from typing import Iterable

import pandas as pd
import pythoncom
import win32com.client
from docx import Document
from docx.document import Document as _Document
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from xml.etree import ElementTree as ET


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
        doc = word.Documents.Open(docx_path, ReadOnly=False, AddToRecentFiles=False)

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


def parse_args() -> argparse.Namespace:
    default_excel, default_word = build_default_paths()

    parser = argparse.ArgumentParser(
        description=(
            "Update a Word .docx file by replacing a target name with a value "
            "taken from an Excel row (sheet 2, column 2 by default)."
        )
    )
    parser.add_argument(
        "--rowNumber",
        type=int,
        required=True,
        help="Excel row number (2 => second row, e.g., B2 for colIndex=1).",
    )
    parser.add_argument(
        "--excel",
        type=str,
        default=default_excel,
        help=f"Path to Excel workbook (default: {default_excel})",
    )
    parser.add_argument(
        "--word",
        type=str,
        default=default_word,
        help=f"Path to Word .docx file (default: {default_word})",
    )
    parser.add_argument(
        "--sheetIndex",
        type=int,
        default=1,
        help="Zero-based sheet index (default: 1 => second sheet).",
    )
    parser.add_argument(
        "--colIndex",
        type=int,
        default=1,
        help="Zero-based column index (default: 1 => second column / B).",
    )
    parser.add_argument(
        "--target",
        type=str,
        default="Когогину С.А.",
        help=(
            "Target string in the Word document to replace. "
            "If omitted, use --startIndex and --length for position mode."
        ),
    )
    parser.add_argument(
        "--startIndex",
        type=int,
        default=None,
        help=(
            "Start position in the concatenated Word text (position mode). "
            "Used only when --target is not provided."
        ),
    )
    parser.add_argument(
        "--length",
        type=int,
        default=None,
        help=(
            "Number of characters to replace from --startIndex (position mode). "
            "Used only when --target is not provided."
        ),
    )
    parser.add_argument(
        "--engine",
        type=str,
        choices=["auto", "word-com", "xml"],
        default="auto",
        help=(
            "Replacement engine for target mode: "
            "auto (prefer Word COM), word-com (force COM), xml (force XML strategy). "
            "Position mode always uses XML strategy."
        ),
    )

    return parser.parse_args()


def main() -> None:
    args = parse_args()

    replacement = read_excel_value(
        excel_path=args.excel,
        sheet_index=args.sheetIndex,
        col_index=args.colIndex,
        row_number=args.rowNumber,
    )

    word_dir, word_name = os.path.split(args.word)
    base, ext = os.path.splitext(word_name)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"{base}.updated.row{args.rowNumber}.{timestamp}{ext}"
    output_path = os.path.join(word_dir, output_name)

    if args.startIndex is not None or args.length is not None:
        if args.startIndex is None or args.length is None:
            raise ValueError("Position mode requires both --startIndex and --length.")
        replace_span_in_docx(
            docx_path=args.word,
            start_index=args.startIndex,
            length=args.length,
            replacement=replacement,
            output_path=output_path,
        )
    else:
        if args.engine in ("auto", "word-com"):
            try:
                replace_text_with_word_com(
                    docx_path=args.word,
                    output_path=output_path,
                    target=args.target,
                    replacement=replacement,
                )
            except Exception:
                if args.engine == "word-com":
                    raise
                replace_text_in_docx(
                    docx_path=args.word,
                    target=args.target,
                    replacement=replacement,
                    output_path=output_path,
                )
        else:
            replace_text_in_docx(
                docx_path=args.word,
                target=args.target,
                replacement=replacement,
                output_path=output_path,
            )

    print("Excel value:", replacement)
    print("Updated Word file written to:", output_path)


if __name__ == "__main__":
    main()


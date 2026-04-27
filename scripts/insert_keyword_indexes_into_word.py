from __future__ import annotations

import argparse
from pathlib import Path
import re
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import clean_keyword, parse_pages, read_csv_rows, sort_keywords

try:
    import win32com.client as win32
except ImportError:
    win32 = None


INPUT_DIR = ROOT_DIR / "input"
OUTPUT_DIR = ROOT_DIR / "output"

DEFAULT_SOURCE_DOCX = INPUT_DIR / "test1.docx"

RU_INDEX_TXT = "keyword_index_ru.txt"
EN_INDEX_TXT = "keyword_index_en.txt"
RU_INDEX_CSV = "keyword_index_ru.csv"
EN_INDEX_CSV = "keyword_index_en.csv"
FORMULA_OPEN_TO_CLOSE = {"(": ")", "[": "]"}
FORMULA_CLOSE_TO_OPEN = {")": "(", "]": "["}
FORMULA_TOKEN_CHARS = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789()[]")


def build_tagged_name(filename: str, run_tag: str) -> str:
    return f"{run_tag}_{filename}" if run_tag else filename


def build_output_docx_name(source_path: Path, run_tag: str) -> str:
    base_name = run_tag or source_path.stem
    return f"{base_name}_with_keyword_indexes.docx"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Insert RU and EN keyword indexes into a Word document via Word COM."
    )
    parser.add_argument(
        "--docx",
        default=str(DEFAULT_SOURCE_DOCX),
        help=f"Path to the input .docx file. Default: {DEFAULT_SOURCE_DOCX}",
    )
    parser.add_argument(
        "--run-tag",
        default="",
        help="Optional tag prefix for keyword index artifacts and output file, e.g. 'test6'.",
    )
    return parser.parse_args(argv)


def read_index_csv(path: Path) -> dict[str, list[int]]:
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    index: dict[str, list[int]] = {}
    for row in read_csv_rows(file_path):
        keyword = clean_keyword(row.get("keyword") or "")
        pages = parse_pages(row.get("pages") or "")
        if keyword and pages:
            index[keyword] = pages
    return index


def read_index_txt(path: Path) -> dict[str, list[int]]:
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    index: dict[str, list[int]] = {}
    with file_path.open("r", encoding="utf-8-sig") as file_obj:
        for raw_line in file_obj:
            line = raw_line.strip()
            if not line:
                continue

            match = re.match(r"^(?P<keyword>.+?)\s+(?P<pages>\d+(?:,\s*\d+)*)$", line)
            if not match:
                continue

            keyword = match.group("keyword")
            pages_text = match.group("pages")
            keyword = clean_keyword(keyword)
            pages = parse_pages(pages_text)
            if keyword and pages:
                index[keyword] = pages
    return index


def resolve_index_input(run_tag: str, txt_name: str, csv_name: str) -> Path:
    tagged_txt_path = OUTPUT_DIR / build_tagged_name(txt_name, run_tag)
    if tagged_txt_path.exists():
        return tagged_txt_path

    tagged_csv_path = OUTPUT_DIR / build_tagged_name(csv_name, run_tag)
    if tagged_csv_path.exists():
        return tagged_csv_path

    raise FileNotFoundError(f"File not found: {tagged_txt_path}")


def read_index(path: Path) -> dict[str, list[int]]:
    if path.suffix.lower() == ".txt":
        return read_index_txt(path)
    return read_index_csv(path)


def choose_output_path(preferred_path: Path) -> Path:
    if not preferred_path.exists():
        return preferred_path

    stem = preferred_path.stem
    suffix = preferred_path.suffix
    parent = preferred_path.parent

    counter = 1
    while True:
        candidate = parent / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def insert_page_break(doc) -> None:
    end_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    end_range.InsertBreak(7)  # wdPageBreak


def insert_formatted_paragraph(
    doc,
    text: str,
    *,
    font_name: str = "Times New Roman",
    font_size: int = 12,
    bold: bool = False,
    italic: bool = False,
) -> None:
    insert_at = doc.Content.End - 1
    paragraph_range = doc.Range(insert_at, insert_at)
    paragraph_range.InsertAfter(text)
    paragraph_range.InsertParagraphAfter()

    formatted_range = doc.Range(insert_at, insert_at + len(text))
    formatted_range.Font.Name = font_name
    formatted_range.Font.Size = font_size
    formatted_range.Font.Bold = bold
    formatted_range.Font.Italic = italic


def insert_heading(doc, text: str, level: int = 1) -> None:
    insert_formatted_paragraph(
        doc,
        text,
        font_name="Times New Roman",
        font_size=14 if level == 1 else 13,
        bold=True,
        italic=False,
    )


def insert_plain_paragraph(doc, text: str, font_name: str = "Times New Roman", font_size: int = 12) -> None:
    insert_formatted_paragraph(
        doc,
        text,
        font_name=font_name,
        font_size=font_size,
        bold=False,
        italic=False,
    )


def find_formula_token_bounds(text: str, digit_index: int) -> tuple[int, int]:
    start = digit_index
    while start > 0 and text[start - 1] in FORMULA_TOKEN_CHARS:
        start -= 1

    end = digit_index + 1
    while end < len(text) and text[end] in FORMULA_TOKEN_CHARS:
        end += 1

    return start, end


def parse_formula_group(token: str, start: int = 0, stop_char: str | None = None) -> tuple[bool, int]:
    index = start
    while index < len(token):
        char = token[index]
        if stop_char is not None and char == stop_char:
            return index > start, index + 1

        if char in FORMULA_OPEN_TO_CLOSE:
            ok, index = parse_formula_group(token, index + 1, FORMULA_OPEN_TO_CLOSE[char])
            if not ok:
                return False, index
            while index < len(token) and token[index].isdigit():
                index += 1
            continue

        if char in FORMULA_CLOSE_TO_OPEN:
            return False, index

        if char.isupper():
            index += 1
            if index < len(token) and token[index].islower():
                index += 1
            while index < len(token) and token[index].isdigit():
                index += 1
            continue

        return False, index

    return stop_char is None and index > start, index


def looks_like_formula_token(token: str) -> bool:
    if not token or not any(char.isdigit() for char in token) or not any(char.isupper() for char in token):
        return False

    ok, index = parse_formula_group(token)
    return ok and index == len(token)


def should_subscript_keyword_digit(text: str, digit_index: int) -> bool:
    if not text[digit_index].isdigit():
        return False

    if digit_index == 0 or text[digit_index - 1] not in FORMULA_TOKEN_CHARS:
        return False

    start, end = find_formula_token_bounds(text, digit_index)
    return looks_like_formula_token(text[start:end])


def build_keyword_entry_text_and_subscripts(keyword: str, page_text: str) -> tuple[str, set[int]]:
    full_text = f"{keyword} {page_text}" if page_text else keyword
    subscript_offsets = {
        index
        for index, char in enumerate(keyword)
        if char.isdigit() and should_subscript_keyword_digit(keyword, index)
    }
    return full_text, subscript_offsets


def insert_keyword_entry_paragraph(
    doc,
    keyword: str,
    page_text: str,
    font_name: str = "Times New Roman",
    font_size: int = 12,
) -> None:
    text, subscript_offsets = build_keyword_entry_text_and_subscripts(keyword, page_text)

    insert_at = doc.Content.End - 1
    paragraph_range = doc.Range(insert_at, insert_at)
    paragraph_range.InsertAfter(text)
    paragraph_range.InsertParagraphAfter()

    formatted_range = doc.Range(insert_at, insert_at + len(text))
    formatted_range.Font.Name = font_name
    formatted_range.Font.Size = font_size
    formatted_range.Font.Bold = False
    formatted_range.Font.Italic = False
    formatted_range.Font.Superscript = False
    formatted_range.Font.Subscript = False

    for offset in sorted(subscript_offsets):
        char_range = doc.Range(insert_at + offset, insert_at + offset + 1)
        char_range.Font.Subscript = True


def insert_index_block(doc, heading: str, index: dict[str, list[int]]) -> None:
    insert_heading(doc, heading, level=1)

    for keyword in sort_keywords(index.keys()):
        pages_sorted = sorted(index[keyword])
        page_text = ", ".join(str(p) for p in pages_sorted)
        insert_keyword_entry_paragraph(doc, keyword, page_text)



def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)

    if win32 is None:
        raise RuntimeError("pywin32 is not installed. Install it with: pip install pywin32")

    source_path = Path(args.docx)
    if not source_path.exists():
        raise FileNotFoundError(f"File not found: {source_path}")
    if source_path.suffix.lower() != ".docx":
        raise ValueError(f"Expected a .docx file: {source_path}")

    output_path = choose_output_path(OUTPUT_DIR / build_output_docx_name(source_path, args.run_tag))
    output_path.parent.mkdir(parents=True, exist_ok=True)

    ru_index_path = resolve_index_input(args.run_tag, RU_INDEX_TXT, RU_INDEX_CSV)
    en_index_path = resolve_index_input(args.run_tag, EN_INDEX_TXT, EN_INDEX_CSV)

    print(f"INPUT DOCX:     {source_path}")
    print(f"RUN TAG:        {args.run_tag or '(none)'}")
    print(f"RU INDEX FILE:  {ru_index_path}")
    print(f"EN INDEX FILE:  {en_index_path}")
    print(f"OUTPUT DOCX:    {output_path}")

    ru_index = read_index(ru_index_path)
    en_index = read_index(en_index_path)

    print(f"RU unique keywords: {len(ru_index)}")
    print(f"EN unique keywords: {len(en_index)}")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    doc = None
    try:
        doc = word.Documents.Open(str(source_path.resolve()))

        insert_page_break(doc)
        insert_index_block(doc, "\u041f\u0440\u0435\u0434\u043c\u0435\u0442\u043d\u044b\u0439 \u0443\u043a\u0430\u0437\u0430\u0442\u0435\u043b\u044c", ru_index)
        insert_plain_paragraph(doc, "")
        insert_index_block(doc, "Keyword Index", en_index)

        save_path = output_path
        while True:
            try:
                doc.SaveAs2(str(save_path.resolve()))
                output_path = save_path
                break
            except Exception:
                if not save_path.exists():
                    raise
                save_path = choose_output_path(save_path)

        print("Done.")
        print(f"Created file: {output_path}")

    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        word.Quit()


if __name__ == "__main__":
    main()

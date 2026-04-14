from __future__ import annotations

from pathlib import Path
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

SOURCE_DOCX = INPUT_DIR / "test1.docx"
OUTPUT_DOCX = OUTPUT_DIR / "test1_with_keyword_indexes.docx"

RU_INDEX_CSV = OUTPUT_DIR / "keyword_index_ru.csv"
EN_INDEX_CSV = OUTPUT_DIR / "keyword_index_en.csv"


def read_index_csv(path: Path) -> dict[str, list[int]]:
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    index: dict[str, list[int]] = {}
    for row in read_csv_rows(file_path):
        keyword = clean_keyword(row.get("keyword") or "")
        pages = parse_pages(row.get("pages") or "")
        if keyword and pages:
            index[keyword] = pages
    return index


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


def insert_index_block(doc, heading: str, index: dict[str, list[int]]) -> None:
    insert_heading(doc, heading, level=1)

    for keyword in sort_keywords(index.keys()):
        pages_sorted = sorted(index[keyword])
        line = f"{keyword} {', '.join(str(p) for p in pages_sorted)}"
        insert_plain_paragraph(doc, line)


def main() -> None:
    if win32 is None:
        raise RuntimeError("Не установлен pywin32. Установите: pip install pywin32")

    source_path = Path(SOURCE_DOCX)
    if not source_path.exists():
        raise FileNotFoundError(f"Файл не найден: {source_path}")

    output_path = Path(OUTPUT_DOCX)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    ru_index = read_index_csv(RU_INDEX_CSV)
    en_index = read_index_csv(EN_INDEX_CSV)

    print(f"RU unique keywords: {len(ru_index)}")
    print(f"EN unique keywords: {len(en_index)}")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    doc = None
    try:
        doc = word.Documents.Open(str(source_path.resolve()))

        insert_page_break(doc)
        insert_index_block(doc, "Предметный указатель", ru_index)
        insert_plain_paragraph(doc, "")
        insert_index_block(doc, "Keyword Index", en_index)

        doc.SaveAs2(str(output_path.resolve()))
        print("Готово.")
        print(f"Создан файл: {output_path}")

    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        word.Quit()


if __name__ == "__main__":
    main()

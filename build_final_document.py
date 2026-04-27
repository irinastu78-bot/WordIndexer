from __future__ import annotations

import argparse
from pathlib import Path

from insert_en_blocks_into_word import save_document_with_fallback
from insert_toc_and_author_into_word import (
    AUTHOR_PROGRESS_STEP,
    TOC_ARTICLE_SPACE_AFTER,
    TOC_INTERNAL_SPACE_AFTER,
    TOC_PROGRESS_STEP,
    insert_toc_entry,
    insert_toc_text_paragraph,
    load_toc_entries,
)
from scripts.insert_keyword_indexes_into_word import (
    insert_heading,
    insert_index_block,
    insert_page_break,
    insert_plain_paragraph,
    read_index_txt,
)

try:
    import win32com.client as win32
except ImportError:
    win32 = None


OUTPUT_DIR = Path("output")
DEFAULT_SOURCE_DOCX = Path("input") / "test1.docx"
WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER = 1
TRAILING_PAGE_BREAK_SCAN_CHARS = 256
MANUAL_PAGE_BREAK = "\x0c"
PARAGRAPH_END_MARK = "\r"
TRAILING_IGNORED_CHARS = {"\r", "\n", "\x07", "\x0b", " ", "\t"}

RU_AUTHOR_HEADING = "\u0410\u0432\u0442\u043e\u0440\u0441\u043a\u0438\u0439 \u0443\u043a\u0430\u0437\u0430\u0442\u0435\u043b\u044c"
EN_AUTHOR_HEADING = "Author Index"
RU_KEYWORD_HEADING = "\u041f\u0440\u0435\u0434\u043c\u0435\u0442\u043d\u044b\u0439 \u0443\u043a\u0430\u0437\u0430\u0442\u0435\u043b\u044c"
EN_KEYWORD_HEADING = "Keyword Index"
TOC_HEADING = "\u041e\u0433\u043b\u0430\u0432\u043b\u0435\u043d\u0438\u0435"

TOC_INDEX_ENTRIES = [
    RU_AUTHOR_HEADING,
    EN_AUTHOR_HEADING,
    RU_KEYWORD_HEADING,
    EN_KEYWORD_HEADING,
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build the final ordered Word document with indexes first and RU TOC last."
    )
    parser.add_argument(
        "--docx",
        dest="docx_path",
        type=Path,
        default=DEFAULT_SOURCE_DOCX,
        help="Path to the source .docx file.",
    )
    parser.add_argument(
        "--run-tag",
        dest="run_tag",
        default="",
        help="Optional prefix for tagged input artifacts and output .docx.",
    )
    return parser.parse_args()


def resolve_output_path(base_name: str, run_tag: str) -> Path:
    clean_run_tag = run_tag.strip()
    file_name = f"{clean_run_tag}_{base_name}" if clean_run_tag else base_name
    return OUTPUT_DIR / file_name


def resolve_output_docx_path(run_tag: str) -> Path:
    clean_run_tag = run_tag.strip()
    if clean_run_tag:
        return OUTPUT_DIR / f"{clean_run_tag}_final_ordered.docx"
    return OUTPUT_DIR / "final_ordered.docx"


def read_nonempty_lines(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    lines: list[str] = []
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = " ".join(raw_line.split()).strip()
        if line:
            lines.append(line)
    return lines


def clean_paragraph_text(value: str) -> str:
    text = (value or "").replace("\r", "").replace("\x07", "")
    return " ".join(text.split()).strip()


def normalize_heading_text(value: str) -> str:
    return clean_paragraph_text(value).casefold()


def iter_paragraphs(document):
    count = int(document.Paragraphs.Count)
    for index in range(1, count + 1):
        paragraph = document.Paragraphs(index)
        yield index, paragraph, clean_paragraph_text(str(paragraph.Range.Text))


def get_page_number(paragraph) -> int:
    page_range = paragraph.Range.Duplicate
    page_range.Collapse(1)
    return int(page_range.Information(WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER))


def find_heading_pages(document) -> dict[str, int]:
    pages: dict[str, int] = {}
    heading_by_key = {normalize_heading_text(heading): heading for heading in TOC_INDEX_ENTRIES}
    remaining = set(TOC_INDEX_ENTRIES)

    for _, paragraph, text in iter_paragraphs(document):
        text_key = normalize_heading_text(text)
        heading = heading_by_key.get(text_key)
        if heading is None or heading not in remaining:
            continue
        pages[heading] = get_page_number(paragraph)
        remaining.remove(heading)
        if not remaining:
            break

    if remaining:
        missing = ", ".join(sorted(remaining))
        print_heading_search_diagnostics(document, remaining)
        raise RuntimeError(f"Index headings not found: {missing}")

    return pages


def print_heading_search_diagnostics(document, missing_headings: set[str]) -> None:
    missing_keys = {normalize_heading_text(heading) for heading in missing_headings}
    interesting: list[tuple[int, str]] = []
    tail: list[tuple[int, str]] = []

    for index, _, text in iter_paragraphs(document):
        if text:
            tail.append((index, text))
            tail = tail[-30:]

        text_key = normalize_heading_text(text)
        if not text_key:
            continue
        if any(missing_key in text_key or text_key in missing_key for missing_key in missing_keys):
            interesting.append((index, text))

    print("[diagnostic] heading search failed")
    print("[diagnostic] nearby heading-like paragraphs:")
    if interesting:
        for index, text in interesting[-20:]:
            print(f"[diagnostic] paragraph {index}: {text}")
    else:
        print("[diagnostic] none")

    print("[diagnostic] last non-empty paragraphs:")
    for index, text in tail:
        print(f"[diagnostic] paragraph {index}: {text}")


def document_ends_with_manual_page_break(document) -> bool:
    tail_start = max(document.Content.Start, document.Content.End - TRAILING_PAGE_BREAK_SCAN_CHARS)
    tail_text = str(document.Range(tail_start, document.Content.End).Text)

    for char in reversed(tail_text):
        if char in TRAILING_IGNORED_CHARS:
            continue
        return char == MANUAL_PAGE_BREAK

    return False


def ensure_new_paragraph_after_trailing_page_break(document) -> None:
    if not document_ends_with_manual_page_break(document):
        return

    insert_at = document.Content.End - 1
    paragraph_range = document.Range(insert_at, insert_at)
    paragraph_range.InsertAfter(PARAGRAPH_END_MARK)


def insert_author_index_block(document, heading: str, lines: list[str], *, add_page_break: bool = True) -> None:
    if add_page_break:
        insert_page_break(document)
    else:
        ensure_new_paragraph_after_trailing_page_break(document)
    insert_heading(document, heading, level=1)
    for index, line in enumerate(lines, start=1):
        insert_plain_paragraph(document, line)
        if index % AUTHOR_PROGRESS_STEP == 0 or index == len(lines):
            print(f"{heading} lines inserted: {index}/{len(lines)}")


def insert_keyword_index_block(document, heading: str, path: Path) -> None:
    index = read_index_txt(path)
    insert_index_block(document, heading, index)
    print(f"{heading} keywords inserted: {len(index)}")


def insert_keyword_index_blocks(document, ru_keyword_path: Path, en_keyword_path: Path) -> None:
    insert_page_break(document)
    insert_keyword_index_block(document, RU_KEYWORD_HEADING, ru_keyword_path)
    insert_plain_paragraph(document, "")
    insert_keyword_index_block(document, EN_KEYWORD_HEADING, en_keyword_path)


def insert_toc_block(document, run_tag: str, index_pages: dict[str, int]) -> None:
    toc_entries = load_toc_entries(run_tag)

    insert_page_break(document)
    insert_heading(document, TOC_HEADING, level=1)
    insert_plain_paragraph(document, "")

    print("inserting RU TOC entries...")
    for index, entry in enumerate(toc_entries, start=1):
        insert_toc_entry(
            document,
            title=str(entry["title"]),
            authors=str(entry["authors"]),
            page=entry["page"] if isinstance(entry["page"], int) else None,
            title_paragraph_index=(
                entry["title_paragraph_index"]
                if isinstance(entry["title_paragraph_index"], int)
                else None
            ),
            is_special=bool(entry.get("is_special")),
        )
        if index % TOC_PROGRESS_STEP == 0 or index == len(toc_entries):
            print(f"TOC entries inserted: {index}/{len(toc_entries)}")

    print("inserting TOC index entries...")
    for index, title in enumerate(TOC_INDEX_ENTRIES, start=1):
        is_last = index == len(TOC_INDEX_ENTRIES)
        insert_toc_text_paragraph(
            document,
            text=title,
            italic=False,
            keep_with_next=not is_last,
            space_after=TOC_ARTICLE_SPACE_AFTER if is_last else TOC_INTERNAL_SPACE_AFTER,
            page_text=str(index_pages[title]),
        )
        print(f"- {title}: {index_pages[title]}")


def main() -> int:
    args = parse_args()
    if win32 is None:
        raise RuntimeError("pywin32 is not installed. Install it with: pip install pywin32")

    run_tag = args.run_tag.strip()
    source_path = Path(args.docx_path).resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"File not found: {source_path}")

    ru_author_path = resolve_output_path("author_index_ru_from_snapshot.txt", run_tag)
    en_author_path = resolve_output_path("draft_author_index_en.txt", run_tag)
    ru_keyword_path = resolve_output_path("keyword_index_ru.txt", run_tag)
    en_keyword_path = resolve_output_path("keyword_index_en.txt", run_tag)
    output_docx_path = resolve_output_docx_path(run_tag)
    output_docx_path.parent.mkdir(parents=True, exist_ok=True)

    ru_author_lines = read_nonempty_lines(ru_author_path)
    en_author_lines = read_nonempty_lines(en_author_path)

    print(f"source docx path: {source_path}")
    print(f"RU author source: {ru_author_path.resolve()}")
    print(f"EN author source: {en_author_path.resolve()}")
    print(f"RU keyword source: {ru_keyword_path.resolve()}")
    print(f"EN keyword source: {en_keyword_path.resolve()}")
    print(f"output docx path: {output_docx_path.resolve()}")
    print("opening Word...")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    document = None
    try:
        document = word.Documents.Open(str(source_path))

        print("inserting RU author index...")
        insert_author_index_block(
            document,
            RU_AUTHOR_HEADING,
            ru_author_lines,
            add_page_break=not document_ends_with_manual_page_break(document),
        )

        print("inserting EN author index...")
        insert_author_index_block(document, EN_AUTHOR_HEADING, en_author_lines)

        print("inserting RU/EN keyword indexes...")
        insert_keyword_index_blocks(document, ru_keyword_path, en_keyword_path)

        print("repaginating before final TOC...")
        document.Repaginate()
        index_pages = find_heading_pages(document)

        print("inserting final RU TOC...")
        insert_toc_block(document, run_tag, index_pages)

        print("saving output docx...")
        saved_output_path = save_document_with_fallback(document, output_docx_path)
        print(f"output docx: {saved_output_path}")
        print("done")
        return 0
    finally:
        if document is not None:
            document.Close(SaveChanges=False)
        word.Quit()


if __name__ == "__main__":
    raise SystemExit(main())

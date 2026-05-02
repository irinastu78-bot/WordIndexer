from __future__ import annotations

import argparse
import csv
import shutil
from pathlib import Path

from insert_toc_and_author_into_word import (
    AUTHOR_PROGRESS_STEP,
    TOC_ARTICLE_SPACE_AFTER,
    TOC_INTERNAL_SPACE_AFTER,
    TOC_PROGRESS_STEP,
    insert_normalized_title_text,
    insert_toc_entry,
    insert_toc_text_paragraph,
    load_toc_entries,
    wrap_author_lines,
)
from scripts.insert_keyword_indexes_into_word import (
    insert_heading,
    insert_index_block,
    insert_page_break,
    insert_plain_paragraph,
    read_index_txt,
    should_subscript_keyword_digit,
)
from scripts.build_draft_toc_en import (
    OUTPUT_CSV as EN_TOC_CSV,
    OUTPUT_TXT as EN_TOC_TXT,
    SOURCE_PAIRS_CSV as EN_TOC_SOURCE_PAIRS_CSV,
    build_rows as build_en_toc_rows,
    build_toc_text as build_en_toc_text,
    safe_write_with_replace as safe_write_en_toc_with_replace,
    write_toc_csv as write_en_toc_csv,
)
from wordkeywords.common import write_text

try:
    import win32com.client as win32
except ImportError:
    win32 = None


OUTPUT_DIR = Path("output")
DEFAULT_SOURCE_DOCX = Path("input") / "test1.docx"
WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER = 1
WD_SECTION_BREAK_NEXT_PAGE = 2
WD_SECTION_BREAK_CONTINUOUS = 3
TRAILING_PAGE_BREAK_SCAN_CHARS = 256
MANUAL_PAGE_BREAK = "\x0c"
PARAGRAPH_END_MARK = "\r"
TRAILING_IGNORED_CHARS = {"\r", "\n", "\x07", "\x0b", " ", "\t"}

RU_AUTHOR_HEADING = "\u0410\u0432\u0442\u043e\u0440\u0441\u043a\u0438\u0439 \u0443\u043a\u0430\u0437\u0430\u0442\u0435\u043b\u044c"
EN_AUTHOR_HEADING = "Author Index"
RU_KEYWORD_HEADING = "\u041f\u0440\u0435\u0434\u043c\u0435\u0442\u043d\u044b\u0439 \u0443\u043a\u0430\u0437\u0430\u0442\u0435\u043b\u044c"
EN_KEYWORD_HEADING = "Keyword Index"
TOC_HEADING = "\u041e\u0433\u043b\u0430\u0432\u043b\u0435\u043d\u0438\u0435"
EN_TOC_HEADING = "Table of Contents"

TOC_INDEX_ENTRIES = [
    RU_AUTHOR_HEADING,
    EN_AUTHOR_HEADING,
    RU_KEYWORD_HEADING,
    EN_KEYWORD_HEADING,
]

REQUIRED_FINAL_ARTIFACTS = [
    ("RU author index", "author_index_ru_from_snapshot.txt"),
    ("EN author index", "draft_author_index_en.txt"),
    ("RU keyword index", "keyword_index_ru.txt"),
    ("EN keyword index", "keyword_index_en.txt"),
    ("RU TOC draft", "draft_toc_ru.csv"),
    ("RU TOC structure debug", "ru_title_paragraph_structure_debug.csv"),
    ("EN TOC draft", EN_TOC_CSV),
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


def is_retryable_copy_error(error: OSError) -> bool:
    if isinstance(error, PermissionError):
        return True
    return getattr(error, "winerror", None) == 32


def build_numbered_docx_path(path: Path, index: int) -> Path:
    return path.with_name(f"{path.stem}_{index}{path.suffix}")


def copy_source_docx_to_output(source_path: Path, output_docx_path: Path) -> Path:
    source = source_path.resolve()
    target = output_docx_path.resolve()
    if source == target:
        raise ValueError(f"Output path must differ from source path: {target}")

    target.parent.mkdir(parents=True, exist_ok=True)
    last_error: OSError | None = None

    for index in range(1000):
        candidate = target if index == 0 else build_numbered_docx_path(target, index)
        try:
            shutil.copy2(source, candidate)
            if candidate != target:
                print(f"fallback output docx: {candidate}")
            return candidate
        except OSError as error:
            if not is_retryable_copy_error(error):
                raise
            last_error = error

    raise RuntimeError(f"Could not prepare output docx near: {target}") from last_error


def build_test_run_commands(docx_path: Path, run_tag: str) -> list[str]:
    docx_text = str(docx_path)
    return [
        f"python dump_doc_paragraph_snapshot.py --docx {docx_text} --run-tag {run_tag}",
        f"python build_author_title_paragraph_ru_from_snapshot.py --run-tag {run_tag}",
        f"python scripts/enrich_ru_title_paragraph_structure.py --docx {docx_text} --run-tag {run_tag}",
        f"python build_draft_author_index_ru_from_snapshot.py --run-tag {run_tag}",
        f"python build_author_index_ru_text_from_snapshot.py --run-tag {run_tag}",
        f"python debug_toc_ru_from_word.py --docx {docx_text} --run-tag {run_tag}",
        f"python build_toc_ru_draft_text.py --run-tag {run_tag}",
        f"python scripts/debug_author_windows_en.py --run-tag {run_tag}",
        f"python scripts/debug_en_title_author_pairs.py --docx {docx_text} --run-tag {run_tag}",
        f"python scripts/build_draft_author_index_en.py --docx {docx_text} --run-tag {run_tag}",
        f"python scripts/build_draft_toc_en.py --docx {docx_text} --run-tag {run_tag}",
        f"python scripts/fast_keyword_index_find.py --docx {docx_text} --run-tag {run_tag}",
        f"python scripts/build_separate_keyword_indexes.py --run-tag {run_tag}",
        f"python build_final_document.py --docx {docx_text} --run-tag {run_tag}",
    ]


def check_required_final_artifacts(run_tag: str, docx_path: Path) -> None:
    missing = [
        (description, resolve_output_path(file_name, run_tag))
        for description, file_name in REQUIRED_FINAL_ARTIFACTS
        if not resolve_output_path(file_name, run_tag).exists()
    ]
    if not missing:
        return

    lines = [
        "Missing prerequisite artifacts for final document build:",
        *[f"- {description}: {path}" for description, path in missing],
        "",
        "Build them for this input/run-tag with:",
        *build_test_run_commands(docx_path, run_tag),
    ]
    raise RuntimeError("\n".join(lines))


def read_nonempty_lines(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    lines: list[str] = []
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = " ".join(raw_line.split()).strip()
        if line:
            lines.append(line)
    return lines


def parse_int(value: str) -> int | None:
    normalized = (value or "").strip()
    if normalized.isdigit():
        return int(normalized)
    return None


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


def get_sections_count(document) -> int:
    try:
        return int(document.Sections.Count)
    except Exception:
        return 0


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


def find_existing_heading_pages(document, headings: list[str]) -> dict[str, int]:
    pages: dict[str, int] = {}
    heading_by_key = {normalize_heading_text(heading): heading for heading in headings}
    remaining = set(headings)

    for _, paragraph, text in iter_paragraphs(document):
        heading = heading_by_key.get(normalize_heading_text(text))
        if heading is None or heading not in remaining:
            continue

        pages[heading] = get_page_number(paragraph)
        remaining.remove(heading)
        if not remaining:
            break

    return pages


def print_heading_page_diagnostics(document) -> None:
    headings = [
        RU_AUTHOR_HEADING,
        EN_AUTHOR_HEADING,
        RU_KEYWORD_HEADING,
        EN_KEYWORD_HEADING,
        TOC_HEADING,
        EN_TOC_HEADING,
    ]
    pages = find_existing_heading_pages(document, headings)
    print("[heading page diagnostic] heading pages:")
    for heading in headings:
        page_text = pages.get(heading, "not-found")
        print(f"[heading page diagnostic] {heading}: {page_text}")


def print_selected_heading_page_diagnostics(document, headings: list[str], label: str) -> None:
    pages = find_existing_heading_pages(document, headings)
    print(f"[heading page diagnostic] {label}:")
    for heading in headings:
        page_text = pages.get(heading, "not-found")
        print(f"[heading page diagnostic] {heading}: {page_text}")


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


def iter_section_header_footers(section):
    for collection_name in ("Headers", "Footers"):
        try:
            collection = getattr(section, collection_name)
        except Exception:
            continue

        for index in range(1, 4):
            try:
                yield collection_name, index, collection(index)
            except Exception:
                continue


def configure_appendix_section(section, *, columns: int) -> None:
    try:
        section.PageSetup.TextColumns.SetCount(columns)
    except Exception as error:
        print(f"[warning] could not set section columns={columns}: {error}")
    try:
        section.PageSetup.DifferentFirstPageHeaderFooter = False
    except Exception as error:
        print(f"[warning] could not disable different first page footer: {error}")


def configure_appendix_sections(document, start_section_index: int, *, columns: int = 1) -> None:
    try:
        sections_count = int(document.Sections.Count)
    except Exception:
        return

    for section_index in range(max(1, start_section_index), sections_count + 1):
        try:
            configure_appendix_section(document.Sections(section_index), columns=columns)
        except Exception as error:
            print(f"[warning] could not configure appendix section {section_index}: {error}")


def repair_appendix_page_numbering(
    document,
    start_section_index: int,
    *,
    appendix_start_page: int,
    columns: int = 1,
) -> None:
    try:
        sections_count = int(document.Sections.Count)
    except Exception:
        return

    print(f"[page numbering repair] appendix_start_page={appendix_start_page}")
    for section_index in range(max(1, start_section_index), sections_count + 1):
        try:
            section = document.Sections(section_index)
        except Exception as error:
            print(f"[warning] could not access appendix section {section_index}: {error}")
            continue

        configure_appendix_section(section, columns=columns)

        for collection_name, header_footer_index, header_footer in iter_section_header_footers(section):
            try:
                header_footer.LinkToPrevious = False
            except Exception as error:
                print(
                    f"[warning] could not unlink {collection_name}({header_footer_index}) "
                    f"for section {section_index}: {error}"
                )
            try:
                header_footer.PageNumbers.RestartNumberingAtSection = True
            except Exception as error:
                print(
                    f"[warning] could not enable page restart for {collection_name}({header_footer_index}) "
                    f"section {section_index}: {error}"
                )
            try:
                header_footer.PageNumbers.StartingNumber = appendix_start_page
            except Exception as error:
                print(
                    f"[warning] could not set page start for {collection_name}({header_footer_index}) "
                    f"section {section_index}: {error}"
                )


def get_header_footer_property(section, collection_name: str, index: int, property_name: str):
    try:
        header_footer = getattr(section, collection_name)(index)
        if property_name == "LinkToPrevious":
            return bool(header_footer.LinkToPrevious)
        if property_name == "RestartNumberingAtSection":
            return bool(header_footer.PageNumbers.RestartNumberingAtSection)
        if property_name == "StartingNumber":
            return int(header_footer.PageNumbers.StartingNumber)
        return "unknown-property"
    except Exception as error:
        return f"unavailable:{error}"


def get_different_first_page_flag(section):
    try:
        return bool(section.PageSetup.DifferentFirstPageHeaderFooter)
    except Exception as error:
        return f"unavailable:{error}"


def get_section_columns_count(section):
    try:
        return int(section.PageSetup.TextColumns.Count)
    except Exception as error:
        return f"unavailable:{error}"


def print_appendix_section_diagnostics(document, start_section_index: int) -> None:
    try:
        sections_count = int(document.Sections.Count)
    except Exception as error:
        print(f"[section diagnostic] unavailable: {error}")
        return

    print("[section diagnostic] appendix sections:")
    print(f"[section diagnostic] appendix_start_section_index={start_section_index}")
    print(f"[section diagnostic] sections_count={sections_count}")
    for section_index in range(max(1, start_section_index), sections_count + 1):
        try:
            section = document.Sections(section_index)
        except Exception as error:
            print(f"[section diagnostic] section {section_index}: unavailable:{error}")
            continue

        columns = get_section_columns_count(section)
        different_first_page = get_different_first_page_flag(section)
        print(
            f"[section diagnostic] section {section_index}: "
            f"columns={columns}; "
            f"different_first_page={different_first_page}"
        )
        for collection_name, label in (("Headers", "Header"), ("Footers", "Footer")):
            for index, type_label in ((1, "Primary"), (2, "FirstPage"), (3, "EvenPages")):
                link_to_previous = get_header_footer_property(
                    section,
                    collection_name,
                    index,
                    "LinkToPrevious",
                )
                restart = get_header_footer_property(
                    section,
                    collection_name,
                    index,
                    "RestartNumberingAtSection",
                )
                starting_number = get_header_footer_property(
                    section,
                    collection_name,
                    index,
                    "StartingNumber",
                )
                print(
                    f"[section diagnostic] section {section_index} {type_label}{label}: "
                    f"link_to_previous={link_to_previous}; "
                    f"restart={restart}; "
                    f"starting_number={starting_number}"
                )


def start_appendix_section(document, *, columns: int = 1, reuse_trailing_page_break: bool = False) -> int:
    break_type = (
        WD_SECTION_BREAK_CONTINUOUS
        if reuse_trailing_page_break and document_ends_with_manual_page_break(document)
        else WD_SECTION_BREAK_NEXT_PAGE
    )
    insert_at = document.Content.End - 1
    section_range = document.Range(insert_at, insert_at)
    section_range.InsertBreak(break_type)
    section_index = int(document.Sections.Count)
    configure_appendix_section(document.Sections(section_index), columns=columns)
    return section_index


def insert_author_index_block(
    document,
    heading: str,
    lines: list[str],
    *,
    reuse_trailing_page_break: bool = False,
    columns: int = 1,
    start_new_section: bool = True,
) -> int:
    if start_new_section:
        section_index = start_appendix_section(
            document,
            columns=columns,
            reuse_trailing_page_break=reuse_trailing_page_break,
        )
    else:
        insert_page_break(document)
        section_index = int(document.Sections.Count)
        configure_appendix_section(document.Sections(section_index), columns=columns)

    insert_heading(document, heading, level=1)
    for index, line in enumerate(lines, start=1):
        insert_plain_paragraph(document, line)
        if index % AUTHOR_PROGRESS_STEP == 0 or index == len(lines):
            print(f"{heading} lines inserted: {index}/{len(lines)}")
    return section_index


def insert_keyword_index_block(document, heading: str, path: Path) -> None:
    index = read_index_txt(path)
    insert_index_block(document, heading, index)
    print(f"{heading} keywords inserted: {len(index)}")


def insert_keyword_index_blocks(document, ru_keyword_path: Path, en_keyword_path: Path) -> None:
    insert_page_break(document)
    configure_appendix_section(document.Sections(document.Sections.Count), columns=1)
    insert_keyword_index_block(document, RU_KEYWORD_HEADING, ru_keyword_path)
    insert_plain_paragraph(document, "")
    insert_keyword_index_block(document, EN_KEYWORD_HEADING, en_keyword_path)


def ensure_en_toc_draft(run_tag: str) -> Path:
    csv_path = resolve_output_path(EN_TOC_CSV, run_tag)
    if csv_path.exists():
        return csv_path

    source_csv_path = resolve_output_path(EN_TOC_SOURCE_PAIRS_CSV, run_tag)
    rows = build_en_toc_rows(source_csv_path)
    txt_path = resolve_output_path(EN_TOC_TXT, run_tag)
    safe_write_en_toc_with_replace(csv_path, lambda output_path: write_en_toc_csv(output_path, rows))
    safe_write_en_toc_with_replace(txt_path, lambda output_path: write_text(output_path, build_en_toc_text(rows)))
    return csv_path


def load_en_toc_entries(run_tag: str) -> list[dict[str, object]]:
    csv_path = ensure_en_toc_draft(run_tag)
    entries: list[dict[str, object]] = []

    with csv_path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh, delimiter=";")
        for row in reader:
            title = (row.get("title") or "").strip()
            page = parse_int(row.get("page") or "")
            if not title or page is None:
                continue
            entries.append(
                {
                    "title": title,
                    "authors": (row.get("authors") or "").strip(),
                    "page": page,
                }
            )

    return entries


def apply_subscripts_to_toc_title(paragraph) -> None:
    paragraph_range = paragraph.Range.Duplicate
    if paragraph_range.End > paragraph_range.Start:
        paragraph_range.End -= 1

    paragraph_text = str(paragraph_range.Text)
    title_text = paragraph_text.split("\t", 1)[0]
    if not title_text:
        return

    for offset, char in enumerate(title_text):
        if not char.isdigit() or not should_subscript_keyword_digit(title_text, offset):
            continue
        char_range = paragraph_range.Duplicate
        char_range.Start = paragraph_range.Start + offset
        char_range.End = char_range.Start + 1
        char_range.Font.Subscript = True


def build_en_toc_title_specs(title: str) -> list[dict[str, object]]:
    normalized = clean_paragraph_text(title)
    return [
        {
            "text": char,
            "superscript": False,
            "subscript": char.isdigit() and should_subscript_keyword_digit(normalized, offset),
        }
        for offset, char in enumerate(normalized)
    ]


def insert_en_toc_entry(document, *, title: str, authors: str, page: int) -> None:
    author_lines = wrap_author_lines(authors)
    page_text = str(page)

    if not author_lines:
        title_paragraph = insert_toc_text_paragraph(
            document,
            text=title,
            italic=False,
            keep_with_next=False,
            space_after=TOC_ARTICLE_SPACE_AFTER,
            page_text=page_text,
        )
        if title_paragraph is not None:
            apply_subscripts_to_toc_title(title_paragraph)
        return

    insert_normalized_title_text(
        document,
        build_en_toc_title_specs(title),
        keep_with_next=True,
        space_after=TOC_INTERNAL_SPACE_AFTER,
    )

    for line_index, author_line in enumerate(author_lines, start=1):
        is_last_line = line_index == len(author_lines)
        insert_toc_text_paragraph(
            document,
            text=author_line,
            italic=True,
            keep_with_next=not is_last_line,
            space_after=TOC_ARTICLE_SPACE_AFTER if is_last_line else TOC_INTERNAL_SPACE_AFTER,
            page_text=page_text if is_last_line else "",
        )


def insert_en_toc_block(document, run_tag: str) -> None:
    toc_entries = load_en_toc_entries(run_tag)

    insert_page_break(document)
    configure_appendix_section(document.Sections(document.Sections.Count), columns=1)
    insert_heading(document, EN_TOC_HEADING, level=1)
    insert_plain_paragraph(document, "")

    print("inserting EN TOC entries...")
    for index, entry in enumerate(toc_entries, start=1):
        insert_en_toc_entry(
            document,
            title=str(entry["title"]),
            authors=str(entry["authors"]),
            page=int(entry["page"]),
        )
        if index % TOC_PROGRESS_STEP == 0 or index == len(toc_entries):
            print(f"EN TOC entries inserted: {index}/{len(toc_entries)}")


def insert_toc_block(document, run_tag: str, index_pages: dict[str, int]) -> None:
    toc_entries = load_toc_entries(run_tag)

    insert_page_break(document)
    configure_appendix_section(document.Sections(document.Sections.Count), columns=1)
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
    check_required_final_artifacts(run_tag, args.docx_path)
    working_docx_path = copy_source_docx_to_output(source_path, output_docx_path)

    ru_author_lines = read_nonempty_lines(ru_author_path)
    en_author_lines = read_nonempty_lines(en_author_path)

    print(f"source docx path: {source_path}")
    print(f"RU author source: {ru_author_path.resolve()}")
    print(f"EN author source: {en_author_path.resolve()}")
    print(f"RU keyword source: {ru_keyword_path.resolve()}")
    print(f"EN keyword source: {en_keyword_path.resolve()}")
    print(f"output docx path: {working_docx_path}")
    print("opening Word...")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    document = None
    try:
        document = word.Documents.Open(str(working_docx_path))
        sections_before_appendices = get_sections_count(document)
        print(f"[section diagnostic] sections before appendices: {sections_before_appendices}")

        print("inserting RU author index...")
        appendix_start_section_index = insert_author_index_block(
            document,
            RU_AUTHOR_HEADING,
            ru_author_lines,
            reuse_trailing_page_break=True,
        )
        print(f"[section diagnostic] appendix_start_section_index={appendix_start_section_index}")

        print("inserting EN author index...")
        insert_author_index_block(
            document,
            EN_AUTHOR_HEADING,
            en_author_lines,
            start_new_section=False,
        )

        print("inserting RU/EN keyword indexes...")
        insert_keyword_index_blocks(document, ru_keyword_path, en_keyword_path)

        configure_appendix_sections(document, appendix_start_section_index, columns=1)
        print(
            "[section diagnostic] sections after index appendices: "
            f"{get_sections_count(document)}"
        )
        print("repaginating before appendix page-number repair...")
        document.Repaginate()
        appendix_start_pages = find_existing_heading_pages(document, [RU_AUTHOR_HEADING])
        appendix_start_page = appendix_start_pages.get(RU_AUTHOR_HEADING)
        if appendix_start_page is None:
            raise RuntimeError(f"Appendix start heading not found: {RU_AUTHOR_HEADING}")
        repair_appendix_page_numbering(
            document,
            appendix_start_section_index,
            appendix_start_page=appendix_start_page,
            columns=1,
        )
        document.Repaginate()
        print_selected_heading_page_diagnostics(
            document,
            [RU_AUTHOR_HEADING, EN_AUTHOR_HEADING, RU_KEYWORD_HEADING, EN_KEYWORD_HEADING],
            "after appendix page-number repair before TOC",
        )

        print("repaginating before final TOC...")
        document.Repaginate()
        index_pages = find_heading_pages(document)

        print("inserting final RU TOC...")
        insert_toc_block(document, run_tag, index_pages)

        print("inserting final EN TOC...")
        insert_en_toc_block(document, run_tag)

        repair_appendix_page_numbering(
            document,
            appendix_start_section_index,
            appendix_start_page=appendix_start_page,
            columns=1,
        )
        document.Repaginate()
        print(f"[section diagnostic] sections after all appendices: {get_sections_count(document)}")
        print_appendix_section_diagnostics(document, appendix_start_section_index)
        print_heading_page_diagnostics(document)

        print("saving output docx...")
        document.Save()
        print(f"output docx: {working_docx_path}")
        print("done")
        return 0
    finally:
        if document is not None:
            document.Close(SaveChanges=False)
        word.Quit()


if __name__ == "__main__":
    raise SystemExit(main())

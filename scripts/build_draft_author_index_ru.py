from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
import re
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import normalize_text, read_csv_rows, write_csv, write_text
from scripts.debug_author_title_paragraph_ru import build_debug_hits


INPUT_DIR = ROOT_DIR / "input"
OUTPUT_DIR = ROOT_DIR / "output"
DOCX_PATH = INPUT_DIR / "test1.docx"

OUTPUT_INDEX_CSV = "draft_author_index_ru.csv"
OUTPUT_INDEX_TXT = "draft_author_index_ru.txt"
OUTPUT_DEBUG_CSV = "draft_author_index_ru_debug.csv"
OUTPUT_DEBUG_TXT = "draft_author_index_ru_debug.txt"

CYR_UPPER = "А-ЯЁ"
CYR_LOWER = "а-яё"
SURNAME_PATTERN = rf"[{CYR_UPPER}][{CYR_LOWER}]+(?:-[{CYR_UPPER}][{CYR_LOWER}]+)?"
COMPOUND_SURNAME_PATTERN = rf"[{CYR_UPPER}][{CYR_LOWER}]+(?:-[{CYR_UPPER}][{CYR_LOWER}]+)+"
GIVEN_NAME_PATTERN = rf"[{CYR_UPPER}][{CYR_LOWER}]+(?:-[{CYR_UPPER}][{CYR_LOWER}]+)?"
INITIAL_PATTERN = rf"[{CYR_UPPER}]\."
OPTIONAL_FINAL_INITIAL_PATTERN = rf"[{CYR_UPPER}]\.?"
DOUBLE_INITIALS_PATTERN = rf"{INITIAL_PATTERN}\s*{OPTIONAL_FINAL_INITIAL_PATTERN}"
SINGLE_INITIAL_PATTERN = OPTIONAL_FINAL_INITIAL_PATTERN

BASE_NAME_PATTERNS: list[tuple[str, str]] = [
    ("compound_full_name", rf"\b{COMPOUND_SURNAME_PATTERN}\s+{GIVEN_NAME_PATTERN}\b"),
    ("surname_initials", rf"\b{SURNAME_PATTERN}\s+{DOUBLE_INITIALS_PATTERN}"),
    ("initials_surname", rf"\b{DOUBLE_INITIALS_PATTERN}\s*{SURNAME_PATTERN}\b"),
    ("surname_single_initial", rf"\b{SURNAME_PATTERN}\s+{SINGLE_INITIAL_PATTERN}"),
]

RAW_AUTHOR_CUE_PATTERNS = [re.compile(pattern) for _, pattern in BASE_NAME_PATTERNS]
NAME_EXTRACTION_PATTERNS = [
    (name, re.compile(pattern))
    for name, pattern in BASE_NAME_PATTERNS
]
AFFILIATION_STRIP_PATTERNS = [
    re.compile(rf"({pattern})(?:\s*\d+(?:\s*,\s*\d+)*)")
    for _, pattern in BASE_NAME_PATTERNS
]


@dataclass
class DraftArticleAuthorRow:
    article_no: int
    article_page: int | None
    article_title: str
    raw_author_line: str
    normalized_author_line: str
    parsed_authors: list[str]
    source_used: str


def split_nonempty_lines(text: str) -> list[str]:
    normalized = normalize_text(text)
    if not normalized:
        return []
    return [line for line in (normalize_text(item) for item in normalized.split("\n")) if line]


def looks_like_author_line(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False
    return any(pattern.search(normalized) for pattern in RAW_AUTHOR_CUE_PATTERNS)


def extract_author_block_from_paragraph(text: str) -> str:
    lines = split_nonempty_lines(text)
    if not lines:
        return ""

    collected: list[str] = []
    for line in reversed(lines):
        if looks_like_author_line(line):
            collected.append(line)
            continue
        if collected:
            break

    collected.reverse()
    return "\n".join(collected)


def normalize_author_name(name: str) -> str:
    normalized = normalize_text(name)
    normalized = re.sub(r"\s+", " ", normalized).strip(" ,;")
    normalized = re.sub(
        rf"\b([{CYR_UPPER}])\.\s*([{CYR_UPPER}])(?![.{CYR_LOWER}])",
        r"\1.\2.",
        normalized,
    )
    normalized = re.sub(rf"\b([{CYR_UPPER}])\.\s*([{CYR_UPPER}])\.", r"\1.\2.", normalized)
    return normalized.strip(" ,;")


def normalize_author_line(text: str) -> str:
    normalized = normalize_text(text)
    if not normalized:
        return ""

    normalized = normalized.replace("*", "")

    for pattern in AFFILIATION_STRIP_PATTERNS:
        normalized = pattern.sub(r"\1", normalized)

    normalized = re.sub(r"(?<=[{upper}{lower}])(?=\d)".format(upper=CYR_UPPER, lower=CYR_LOWER), " ", normalized)
    normalized = re.sub(r"\b\d+(?:\s*,\s*\d+)*\b", "", normalized)
    normalized = re.sub(r"\s+,", ",", normalized)
    normalized = re.sub(r",\s*,+", ", ", normalized)
    normalized = re.sub(r"\s*;\s*", "; ", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    normalized = re.sub(r"\s*\n\s*", "\n", normalized)

    cleaned_lines: list[str] = []
    for line in normalized.split("\n"):
        line = normalize_text(line)
        line = line.strip(" ,;")
        if line:
            cleaned_lines.append(line)

    return "\n".join(cleaned_lines)


def extract_authors_from_line(text: str) -> list[str]:
    normalized = normalize_text(text)
    if not normalized:
        return []

    candidates: list[tuple[int, int, int, str]] = []

    for priority, (_, pattern) in enumerate(NAME_EXTRACTION_PATTERNS):
        for match in pattern.finditer(normalized):
            matched_text = normalize_author_name(match.group(0))
            if matched_text:
                candidates.append((match.start(), match.end(), priority, matched_text))

    candidates.sort(key=lambda item: (item[0], -(item[1] - item[0]), item[2]))

    selected: list[tuple[int, int, str]] = []
    seen: set[str] = set()

    for start, end, _, matched_text in candidates:
        if any(start < used_end and end > used_start for used_start, used_end, _ in selected):
            continue

        key = matched_text.casefold()
        if key in seen:
            continue

        seen.add(key)
        selected.append((start, end, matched_text))

    selected.sort(key=lambda item: item[0])
    return [name for _, _, name in selected]


def extract_authors_from_block(text: str) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()

    for line in split_nonempty_lines(text):
        for name in extract_authors_from_line(line):
            key = name.casefold()
            if key not in seen:
                seen.add(key)
                result.append(name)

    return result


def strip_trailing_author_block(paragraph_text: str, author_block: str) -> str:
    paragraph_lines = split_nonempty_lines(paragraph_text)
    author_lines = split_nonempty_lines(author_block)

    if author_lines and paragraph_lines[-len(author_lines):] == author_lines:
        paragraph_lines = paragraph_lines[:-len(author_lines)]

    return "\n".join(paragraph_lines).strip()


def reparse_row(
    *,
    article_no: int,
    article_page: int | None,
    article_title: str,
    raw_author_line: str,
    source_used: str,
) -> DraftArticleAuthorRow:
    normalized_author_line = normalize_author_line(raw_author_line)
    parsed_authors = extract_authors_from_block(normalized_author_line)

    if not parsed_authors and source_used != "not_found":
        source_used = "not_found"

    return DraftArticleAuthorRow(
        article_no=article_no,
        article_page=article_page,
        article_title=article_title,
        raw_author_line=raw_author_line,
        normalized_author_line=normalized_author_line,
        parsed_authors=parsed_authors,
        source_used=source_used,
    )


def build_draft_rows(docx_path: str) -> list[DraftArticleAuthorRow]:
    hits = build_debug_hits(docx_path)
    rows: list[DraftArticleAuthorRow] = []

    for hit in hits:
        title_author_block = extract_author_block_from_paragraph(hit.title_paragraph_text)
        next_author_block = extract_author_block_from_paragraph(hit.next_paragraph_text)
        article_title = strip_trailing_author_block(hit.title_paragraph_text, title_author_block) or hit.title_paragraph_text

        title_row = reparse_row(
            article_no=hit.article_no,
            article_page=hit.page,
            article_title=article_title,
            raw_author_line=title_author_block,
            source_used="title_paragraph",
        )

        if title_row.parsed_authors:
            rows.append(title_row)
            continue

        next_row = reparse_row(
            article_no=hit.article_no,
            article_page=hit.page,
            article_title=article_title,
            raw_author_line=next_author_block,
            source_used="next_paragraph",
        )

        if next_row.parsed_authors:
            rows.append(next_row)
            continue

        rows.append(
            DraftArticleAuthorRow(
                article_no=hit.article_no,
                article_page=hit.page,
                article_title=article_title,
                raw_author_line=title_author_block or next_author_block,
                normalized_author_line=title_row.normalized_author_line or next_row.normalized_author_line,
                parsed_authors=[],
                source_used="not_found",
            )
        )

    return rows


def build_draft_rows_from_debug_csv(path: Path) -> list[DraftArticleAuthorRow]:
    rows: list[DraftArticleAuthorRow] = []

    for item in read_csv_rows(path):
        article_no_text = normalize_text(item.get("article_no", ""))
        article_page_text = normalize_text(item.get("article_page", ""))
        article_no = int(article_no_text) if article_no_text.isdigit() else 0
        article_page = int(article_page_text) if article_page_text.isdigit() else None
        article_title = normalize_text(item.get("article_title", ""))
        raw_author_line = normalize_text(item.get("raw_author_line", "")) or normalize_text(item.get("normalized_author_line", ""))
        source_used = normalize_text(item.get("source_used", "")) or "not_found"

        rows.append(
            reparse_row(
                article_no=article_no,
                article_page=article_page,
                article_title=article_title,
                raw_author_line=raw_author_line,
                source_used=source_used,
            )
        )

    return rows


def build_author_index(rows: list[DraftArticleAuthorRow]) -> dict[str, list[int]]:
    index: dict[str, set[int]] = defaultdict(set)

    for row in rows:
        if row.article_page is None:
            continue
        for author in row.parsed_authors:
            index[author].add(row.article_page)

    return {
        author: sorted(pages)
        for author, pages in sorted(index.items(), key=lambda item: item[0].casefold())
    }


def format_pages(pages: list[int]) -> str:
    return ", ".join(str(page) for page in pages)


def write_index_csv(path: Path, index: dict[str, list[int]]) -> None:
    rows = ([author, format_pages(pages)] for author, pages in index.items())
    write_csv(path, ["author", "pages"], rows)


def write_debug_csv(path: Path, rows: list[DraftArticleAuthorRow]) -> None:
    csv_rows = (
        [
            row.article_no,
            row.article_page if row.article_page is not None else "",
            row.article_title,
            row.raw_author_line,
            row.normalized_author_line,
            " | ".join(row.parsed_authors),
            row.source_used,
        ]
        for row in rows
    )
    write_csv(
        path,
        [
            "article_no",
            "article_page",
            "article_title",
            "raw_author_line",
            "normalized_author_line",
            "parsed_authors",
            "source_used",
        ],
        csv_rows,
    )


def build_index_text(index: dict[str, list[int]]) -> str:
    lines = [f"{author} -> {format_pages(pages)}" for author, pages in index.items()]
    return "\n".join(lines) + ("\n" if lines else "")


def build_debug_text(rows: list[DraftArticleAuthorRow]) -> str:
    blocks: list[str] = []

    for row in rows:
        page_text = str(row.article_page) if row.article_page is not None else "?"
        parsed_authors = " | ".join(row.parsed_authors) if row.parsed_authors else "<not found>"
        blocks.append(
            "\n".join(
                [
                    "=" * 100,
                    f"ARTICLE {row.article_no} | page {page_text}",
                    "TITLE:",
                    row.article_title or "<empty>",
                    "",
                    "RAW AUTHOR LINE:",
                    row.raw_author_line or "<empty>",
                    "",
                    "NORMALIZED AUTHOR LINE:",
                    row.normalized_author_line or "<empty>",
                    "",
                    f"PARSED AUTHORS: {parsed_authors}",
                    f"SOURCE USED: {row.source_used}",
                ]
            )
        )

    return "\n\n".join(blocks) + ("\n" if blocks else "")


def print_summary(rows: list[DraftArticleAuthorRow], index: dict[str, list[int]]) -> None:
    title_count = sum(1 for row in rows if row.source_used == "title_paragraph")
    next_count = sum(1 for row in rows if row.source_used == "next_paragraph")
    not_found_count = sum(1 for row in rows if row.source_used == "not_found")

    print("=" * 100)
    print(f"RU ARTICLES PROCESSED:  {len(rows)}")
    print(f"TITLE PARAGRAPH USED:   {title_count}")
    print(f"NEXT PARAGRAPH USED:    {next_count}")
    print(f"NOT FOUND:              {not_found_count}")
    print(f"AUTHORS IN INDEX:       {len(index)}")
    print("=" * 100)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--from-debug-csv",
        nargs="?",
        const=str(OUTPUT_DIR / OUTPUT_DEBUG_CSV),
        default="",
        help="Rebuild draft files from an existing debug CSV without Word COM.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    docx_path = Path(DOCX_PATH)
    if not docx_path.exists():
        raise FileNotFoundError(f"Файл не найден: {docx_path}")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.from_debug_csv:
        debug_csv_source = Path(args.from_debug_csv)
        if not debug_csv_source.exists():
            raise FileNotFoundError(f"Файл не найден: {debug_csv_source}")
        rows = build_draft_rows_from_debug_csv(debug_csv_source)
    else:
        rows = build_draft_rows(str(docx_path))
    index = build_author_index(rows)

    index_csv_path = output_dir / OUTPUT_INDEX_CSV
    index_txt_path = output_dir / OUTPUT_INDEX_TXT
    debug_csv_path = output_dir / OUTPUT_DEBUG_CSV
    debug_txt_path = output_dir / OUTPUT_DEBUG_TXT

    write_index_csv(index_csv_path, index)
    write_text(index_txt_path, build_index_text(index))
    write_debug_csv(debug_csv_path, rows)
    write_text(debug_txt_path, build_debug_text(rows))
    print_summary(rows, index)

    print("\nФайлы сохранены:")
    print(index_csv_path)
    print(index_txt_path)
    print(debug_csv_path)
    print(debug_txt_path)


if __name__ == "__main__":
    main()

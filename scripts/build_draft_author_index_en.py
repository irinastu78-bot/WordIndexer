from __future__ import annotations

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


OUTPUT_DIR = ROOT_DIR / "output"
INPUT_DEBUG_CSV = OUTPUT_DIR / "author_block_en_debug.csv"

OUTPUT_INDEX_CSV = "draft_author_index_en.csv"
OUTPUT_INDEX_TXT = "draft_author_index_en.txt"
OUTPUT_DEBUG_CSV = "draft_author_index_en_debug.csv"
OUTPUT_DEBUG_TXT = "draft_author_index_en_debug.txt"

LATIN_NAME = r"[A-Z][A-Za-z'-]+"
INITIAL = r"[A-Z][a-z]{0,2}\."
DOUBLE_INITIALS = rf"{INITIAL}\s*{INITIAL}"

BASE_NAME_PATTERNS: list[tuple[str, str]] = [
    ("surname_double_initials", rf"\b{LATIN_NAME}\s+{DOUBLE_INITIALS}\b"),
    ("initials_surname", rf"\b{DOUBLE_INITIALS}\s*{LATIN_NAME}\b"),
    ("surname_single_initial", rf"\b{LATIN_NAME}\s+{INITIAL}\b"),
    ("initial_surname", rf"\b{INITIAL}\s*{LATIN_NAME}\b"),
]

NAME_EXTRACTION_PATTERNS = [
    (name, re.compile(pattern))
    for name, pattern in BASE_NAME_PATTERNS
]
AFFILIATION_STRIP_PATTERNS = [
    re.compile(rf"({pattern})(?:\s*\d+(?:\s*,\s*\d+)*)")
    for _, pattern in BASE_NAME_PATTERNS
]


@dataclass
class DraftEnAuthorRow:
    article_no: int
    article_page: int | None
    chosen_source: str
    chosen_title_source: str
    en_title_candidate: str
    raw_en_author_line_used: str
    normalized_en_author_line: str
    parsed_en_authors: list[str]
    status: str


def split_nonempty_lines(text: str) -> list[str]:
    normalized = normalize_text(text)
    if not normalized:
        return []
    return [line for line in (normalize_text(item) for item in normalized.split("\n")) if line]


def normalize_author_name(name: str) -> str:
    normalized = normalize_text(name)
    normalized = re.sub(r"\s+", " ", normalized).strip(" ,;")
    normalized = re.sub(r"\b([A-Z])\.\s*([A-Z][a-z]{0,2}\.)", r"\1.\2", normalized)

    initials_surname_match = re.fullmatch(rf"({DOUBLE_INITIALS})\s*({LATIN_NAME})", normalized)
    if initials_surname_match:
        initials = normalize_text(initials_surname_match.group(1)).replace(" ", "")
        surname = initials_surname_match.group(2)
        return f"{surname} {initials}".strip(" ,;")

    initial_surname_match = re.fullmatch(rf"({INITIAL})\s*({LATIN_NAME})", normalized)
    if initial_surname_match:
        initial = normalize_text(initial_surname_match.group(1)).replace(" ", "")
        surname = initial_surname_match.group(2)
        return f"{surname} {initial}".strip(" ,;")

    surname_initials_match = re.fullmatch(rf"({LATIN_NAME})\s+({DOUBLE_INITIALS})", normalized)
    if surname_initials_match:
        surname = surname_initials_match.group(1)
        initials = normalize_text(surname_initials_match.group(2)).replace(" ", "")
        return f"{surname} {initials}".strip(" ,;")

    surname_initial_match = re.fullmatch(rf"({LATIN_NAME})\s+({INITIAL})", normalized)
    if surname_initial_match:
        surname = surname_initial_match.group(1)
        initial = normalize_text(surname_initial_match.group(2)).replace(" ", "")
        return f"{surname} {initial}".strip(" ,;")

    return normalized.strip(" ,;")


def normalize_author_line(text: str) -> str:
    normalized = normalize_text(text)
    if not normalized:
        return ""

    normalized = normalized.replace("*", "")

    for pattern in AFFILIATION_STRIP_PATTERNS:
        normalized = pattern.sub(r"\1", normalized)

    normalized = re.sub(r"(?<=[A-Za-z])(?=\d)", " ", normalized)
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


def choose_source(item: dict[str, str]) -> tuple[str, str, str, str]:
    font14_status = normalize_text(item.get("font14_status", ""))
    status = normalize_text(item.get("status", ""))

    if font14_status == "found_title_and_author":
        return (
            "font14",
            "font14",
            normalize_text(item.get("font14_en_title", "")),
            normalize_text(item.get("font14_raw_en_author_line", "")),
        )

    if status == "found_title_and_author":
        return (
            "fallback",
            "fallback",
            normalize_text(item.get("en_title", "")),
            normalize_text(item.get("raw_en_author_line", "")),
        )

    return (
        "not_found",
        "none",
        normalize_text(item.get("font14_en_title", "")) or normalize_text(item.get("en_title", "")),
        "",
    )


def build_rows_from_debug_csv(path: Path) -> list[DraftEnAuthorRow]:
    rows: list[DraftEnAuthorRow] = []

    for item in read_csv_rows(path):
        article_no_text = normalize_text(item.get("article_no", ""))
        article_page_text = normalize_text(item.get("article_page", ""))
        article_no = int(article_no_text) if article_no_text.isdigit() else 0
        article_page = int(article_page_text) if article_page_text.isdigit() else None

        chosen_source, chosen_title_source, en_title_candidate, raw_line = choose_source(item)
        normalized_line = normalize_author_line(raw_line)
        parsed_authors = extract_authors_from_block(normalized_line)
        status = "found" if parsed_authors else "not_found"

        if chosen_source == "not_found" or not parsed_authors:
            chosen_source = "not_found" if chosen_source == "not_found" else chosen_source

        rows.append(
            DraftEnAuthorRow(
                article_no=article_no,
                article_page=article_page,
                chosen_source=chosen_source if parsed_authors or chosen_source == "not_found" else chosen_source,
                chosen_title_source=chosen_title_source,
                en_title_candidate=en_title_candidate,
                raw_en_author_line_used=raw_line,
                normalized_en_author_line=normalized_line,
                parsed_en_authors=parsed_authors,
                status=status,
            )
        )

    return rows


def build_author_index(rows: list[DraftEnAuthorRow]) -> dict[str, list[int]]:
    index: dict[str, set[int]] = defaultdict(set)

    for row in rows:
        if row.article_page is None:
            continue
        for author in row.parsed_en_authors:
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


def write_debug_csv(path: Path, rows: list[DraftEnAuthorRow]) -> None:
    csv_rows = (
        [
            row.article_no,
            row.article_page if row.article_page is not None else "",
            row.chosen_source,
            row.chosen_title_source,
            row.en_title_candidate,
            row.raw_en_author_line_used,
            row.normalized_en_author_line,
            " | ".join(row.parsed_en_authors),
            row.status,
        ]
        for row in rows
    )
    write_csv(
        path,
        [
            "article_no",
            "article_page",
            "chosen_source",
            "chosen_title_source",
            "en_title_candidate",
            "raw_en_author_line_used",
            "normalized_en_author_line",
            "parsed_en_authors",
            "status",
        ],
        csv_rows,
    )


def build_index_text(index: dict[str, list[int]]) -> str:
    lines = [f"{author} -> {format_pages(pages)}" for author, pages in index.items()]
    return "\n".join(lines) + ("\n" if lines else "")


def build_debug_text(rows: list[DraftEnAuthorRow]) -> str:
    blocks: list[str] = []

    for row in rows:
        page_text = str(row.article_page) if row.article_page is not None else "?"
        parsed_authors = " | ".join(row.parsed_en_authors) if row.parsed_en_authors else "<not found>"
        blocks.append(
            "\n".join(
                [
                    "=" * 100,
                    f"ARTICLE {row.article_no} | page {page_text}",
                    f"CHOSEN SOURCE: {row.chosen_source}",
                    f"CHOSEN TITLE SOURCE: {row.chosen_title_source}",
                    f"STATUS: {row.status}",
                    "",
                    "EN TITLE CANDIDATE:",
                    row.en_title_candidate or "<empty>",
                    "",
                    "RAW EN AUTHOR LINE USED:",
                    row.raw_en_author_line_used or "<empty>",
                    "",
                    "NORMALIZED EN AUTHOR LINE:",
                    row.normalized_en_author_line or "<empty>",
                    "",
                    f"PARSED EN AUTHORS: {parsed_authors}",
                ]
            )
        )

    return "\n\n".join(blocks) + ("\n" if blocks else "")


def print_summary(rows: list[DraftEnAuthorRow], index: dict[str, list[int]]) -> None:
    font14_count = sum(1 for row in rows if row.chosen_source == "font14")
    fallback_count = sum(1 for row in rows if row.chosen_source == "fallback")
    not_found_count = sum(1 for row in rows if row.chosen_source == "not_found")

    print("=" * 100)
    print(f"EN ARTICLES PROCESSED:   {len(rows)}")
    print(f"FONT14 USED:             {font14_count}")
    print(f"FALLBACK USED:           {fallback_count}")
    print(f"NOT FOUND:               {not_found_count}")
    print(f"AUTHORS IN INDEX:        {len(index)}")
    print("=" * 100)


def main() -> None:
    input_path = Path(INPUT_DEBUG_CSV)
    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    rows = build_rows_from_debug_csv(input_path)
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

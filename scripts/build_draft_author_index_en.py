from __future__ import annotations

import argparse
import os
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
import re
import sys
import uuid

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import normalize_text, write_csv, write_text
from scripts.debug_en_title_author_pairs import (
    DOCX_PATH,
    EnTitleAuthorPairRow,
    build_rows as build_pair_rows,
)


OUTPUT_DIR = ROOT_DIR / "output"
SOURCE_WINDOWS_CSV = "author_windows_en_debug.csv"

OUTPUT_INDEX_CSV = "draft_author_index_en.csv"
OUTPUT_INDEX_TXT = "draft_author_index_en.txt"
OUTPUT_DEBUG_CSV = "draft_author_index_en_debug.csv"
OUTPUT_DEBUG_TXT = "draft_author_index_en_debug.txt"

LATIN_NAME = r"[A-Z][A-Za-z'-]+"
INITIAL_LETTER = r"(?:[A-Z]|[\u0410-\u042f\u0401\u0451])"
INITIAL_TAIL = r"(?:[a-z]|[\u0430-\u044f\u0451]){0,2}"
INITIAL = rf"{INITIAL_LETTER}{INITIAL_TAIL}\s*\."
DOUBLE_INITIALS = rf"{INITIAL}\s*{INITIAL}"
NAME_END = r"(?=[\W\d]|$)"

BASE_NAME_PATTERNS: list[tuple[str, str]] = [
    ("surname_double_initials", rf"\b{LATIN_NAME}\s+{DOUBLE_INITIALS}{NAME_END}"),
    ("initials_surname", rf"\b{DOUBLE_INITIALS}\s*{LATIN_NAME}{NAME_END}"),
    ("surname_single_initial", rf"\b{LATIN_NAME}\s+{INITIAL}{NAME_END}"),
    ("initial_surname", rf"\b{INITIAL}\s*{LATIN_NAME}{NAME_END}"),
]

NAME_EXTRACTION_PATTERNS = [
    (name, re.compile(pattern))
    for name, pattern in BASE_NAME_PATTERNS
]
AFFILIATION_STRIP_PATTERNS = [
    re.compile(rf"({pattern})(?:\s*\d+(?:\s*,\s*\d+)*)")
    for _, pattern in BASE_NAME_PATTERNS
]
CYRILLIC_HOMOGLYPH_MAP = str.maketrans({
    "А": "A",
    "В": "B",
    "С": "C",
    "Е": "E",
    "К": "K",
    "М": "M",
    "Н": "H",
    "О": "O",
    "Р": "P",
    "Т": "T",
    "Х": "X",
    "У": "Y",
})
TITLE_SERVICE_CUES = (
    "student",
    "graduate student",
    "junior researcher",
    "masters student",
    "master's student",
    "masters's student",
)


@dataclass
class DraftEnAuthorRow:
    article_no: int
    article_page: int | None
    method: str
    en_title_candidate: str
    raw_en_author_line_used: str
    normalized_en_author_line: str
    parsed_en_authors: list[str]
    extraction_status: str
    status: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build the draft EN author index directly from the source .docx."
    )
    parser.add_argument(
        "--docx",
        type=Path,
        default=DOCX_PATH,
        help="Path to source .docx. Defaults to input/test1.docx.",
    )
    parser.add_argument(
        "--run-tag",
        dest="run_tag",
        default="",
        help="Optional prefix for tagged input/output artifact names.",
    )
    return parser.parse_args()


def resolve_output_path(base_name: str, run_tag: str) -> Path:
    clean_run_tag = run_tag.strip()
    file_name = f"{clean_run_tag}_{base_name}" if clean_run_tag else base_name
    return OUTPUT_DIR / file_name


def resolve_source_run_tag(run_tag: str, docx_path: Path) -> str:
    clean_run_tag = run_tag.strip()
    if clean_run_tag:
        return clean_run_tag

    inferred_run_tag = normalize_text(docx_path.stem)
    if inferred_run_tag and resolve_output_path(SOURCE_WINDOWS_CSV, inferred_run_tag).exists():
        print(f"[info] using source artifacts from run-tag '{inferred_run_tag}'")
        return inferred_run_tag

    return ""


def split_nonempty_lines(text: str) -> list[str]:
    normalized = normalize_text(text)
    if not normalized:
        return []
    return [line for line in (normalize_text(item) for item in normalized.split("\n")) if line]


def normalize_author_name(name: str) -> str:
    normalized = normalize_text(name).translate(CYRILLIC_HOMOGLYPH_MAP)
    normalized = re.sub(r"\s+", " ", normalized).strip(" ,;")
    normalized = re.sub(rf"\b({INITIAL_LETTER})\.\s*({INITIAL_LETTER}{INITIAL_TAIL}\.)", r"\1.\2", normalized)

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
    normalized = normalize_text(text).translate(CYRILLIC_HOMOGLYPH_MAP)
    if not normalized:
        return ""

    normalized = normalized.replace("*", "")

    for pattern in AFFILIATION_STRIP_PATTERNS:
        normalized = pattern.sub(r"\1", normalized)

    normalized = re.sub(r"(?<=[A-Za-z\u0410-\u042f\u0401\u0451\u0430-\u044f])(?=\d)", " ", normalized)
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


def parse_author_name_parts(name: str) -> tuple[str, str, str | None] | None:
    normalized = normalize_author_name(name)
    double_match = re.fullmatch(rf"({LATIN_NAME})\s+({INITIAL})({INITIAL})", normalized)
    if double_match:
        return double_match.group(1), double_match.group(2), double_match.group(3)

    single_match = re.fullmatch(rf"({LATIN_NAME})\s+({INITIAL})", normalized)
    if single_match:
        return single_match.group(1), single_match.group(2), None

    return None


def merge_single_initial_variants(rows: list[DraftEnAuthorRow]) -> None:
    fuller_variants: dict[tuple[str, str], set[str]] = defaultdict(set)

    for row in rows:
        for author in row.parsed_en_authors:
            parsed = parse_author_name_parts(author)
            if parsed is None:
                continue
            surname, first_initial, second_initial = parsed
            if second_initial:
                fuller_variants[(surname.casefold(), first_initial.casefold())].add(
                    normalize_author_name(author)
                )

    if not fuller_variants:
        return

    for row in rows:
        merged_authors: list[str] = []
        seen: set[str] = set()

        for author in row.parsed_en_authors:
            resolved_author = normalize_author_name(author)
            parsed = parse_author_name_parts(resolved_author)
            if parsed is not None:
                surname, first_initial, second_initial = parsed
                if second_initial is None:
                    candidates = fuller_variants.get((surname.casefold(), first_initial.casefold()), set())
                    if len(candidates) == 1:
                        resolved_author = next(iter(candidates))

            key = resolved_author.casefold()
            if key not in seen:
                seen.add(key)
                merged_authors.append(resolved_author)

        row.parsed_en_authors = merged_authors


def looks_like_service_title(text: str) -> bool:
    normalized = normalize_text(text).replace("’", "'").replace("*", "").strip(" ,;").casefold()
    if not normalized:
        return False
    return any(cue in normalized for cue in TITLE_SERVICE_CUES)


def build_rows_from_pair_rows(pair_rows: list[EnTitleAuthorPairRow]) -> list[DraftEnAuthorRow]:
    rows: list[DraftEnAuthorRow] = []

    for item in pair_rows:
        extraction_status = normalize_text(item.status)
        method = normalize_text(item.source_used) or "not_found"
        title_candidate = normalize_text(item.en_title_candidate)
        raw_line = normalize_text(item.en_author_candidate)
        if looks_like_service_title(title_candidate):
            title_candidate = ""
            if not raw_line:
                method = "not_found"
                extraction_status = "not_found"
        normalized_line = normalize_author_line(raw_line)
        parsed_authors = extract_authors_from_block(normalized_line)

        if parsed_authors:
            status = "found"
        elif raw_line:
            status = "parse_failed"
        else:
            status = extraction_status or "not_found"

        rows.append(
            DraftEnAuthorRow(
                article_no=item.article_no,
                article_page=item.article_page,
                method=method,
                en_title_candidate=title_candidate,
                raw_en_author_line_used=raw_line,
                normalized_en_author_line=normalized_line,
                parsed_en_authors=parsed_authors,
                extraction_status=extraction_status,
                status=status,
            )
        )

    merge_single_initial_variants(rows)
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
            row.en_title_candidate,
            row.raw_en_author_line_used,
            row.normalized_en_author_line,
            " | ".join(row.parsed_en_authors),
            row.method,
            row.extraction_status,
            row.status,
        ]
        for row in rows
    )
    write_csv(
        path,
        [
            "article_no",
            "page",
            "title",
            "authors_raw",
            "authors_normalized",
            "parsed_authors",
            "method",
            "extraction_status",
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
                    f"METHOD: {row.method}",
                    f"EXTRACTION STATUS: {row.extraction_status or '?'}",
                    f"STATUS: {row.status}",
                    "",
                    "TITLE:",
                    row.en_title_candidate or "<empty>",
                    "",
                    "AUTHORS RAW:",
                    row.raw_en_author_line_used or "<empty>",
                    "",
                    "AUTHORS NORMALIZED:",
                    row.normalized_en_author_line or "<empty>",
                    "",
                    f"PARSED AUTHORS: {parsed_authors}",
                ]
            )
        )

    return "\n\n".join(blocks) + ("\n" if blocks else "")


def print_summary(rows: list[DraftEnAuthorRow], index: dict[str, list[int]]) -> None:
    font14_count = sum(1 for row in rows if row.status == "found" and row.method == "font14")
    fallback_count = sum(1 for row in rows if row.status == "found" and row.method != "font14")
    not_found_count = sum(1 for row in rows if row.status != "found")

    print("=" * 100)
    print(f"EN ARTICLES PROCESSED:   {len(rows)}")
    print(f"FONT14 USED:             {font14_count}")
    print(f"FALLBACK USED:           {fallback_count}")
    print(f"NOT FOUND:               {not_found_count}")
    print(f"AUTHORS IN INDEX:        {len(index)}")
    print("=" * 100)


def is_retryable_output_write_error(error: OSError) -> bool:
    if isinstance(error, PermissionError):
        return True
    return getattr(error, "winerror", None) == 32


def build_fallback_output_path(path: Path, suffix_no: int) -> Path:
    return path.with_name(f"{path.stem}_{suffix_no}{path.suffix}")


def build_temp_output_path(path: Path) -> Path:
    return path.with_name(f"{path.name}.tmp.{os.getpid()}.{uuid.uuid4().hex}")


def safe_write_with_replace(path: Path, writer) -> Path:
    temp_path = build_temp_output_path(path)

    try:
        writer(temp_path)
        os.replace(temp_path, path)
        return path
    except OSError as error:
        try:
            if temp_path.exists():
                temp_path.unlink()
        except OSError:
            pass
        if not is_retryable_output_write_error(error):
            raise

    for suffix_no in range(1, 1000):
        fallback_path = build_fallback_output_path(path, suffix_no)
        temp_fallback_path = build_temp_output_path(fallback_path)
        try:
            writer(temp_fallback_path)
            os.replace(temp_fallback_path, fallback_path)
            print(f"[save] primary path unavailable, saved fallback artifact: {fallback_path}")
            return fallback_path
        except OSError as error:
            try:
                if temp_fallback_path.exists():
                    temp_fallback_path.unlink()
            except OSError:
                pass
            if not is_retryable_output_write_error(error):
                raise

    raise RuntimeError(f"Could not save artifact with fallback names near: {path}")


def main() -> None:
    args = parse_args()
    run_tag = args.run_tag.strip()
    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"File not found: {docx_path}")
    source_run_tag = resolve_source_run_tag(run_tag, docx_path)

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    pair_rows = build_pair_rows(docx_path, source_run_tag)
    rows = build_rows_from_pair_rows(pair_rows)
    index = build_author_index(rows)

    index_csv_path = resolve_output_path(OUTPUT_INDEX_CSV, run_tag)
    index_txt_path = resolve_output_path(OUTPUT_INDEX_TXT, run_tag)
    debug_csv_path = resolve_output_path(OUTPUT_DEBUG_CSV, run_tag)
    debug_txt_path = resolve_output_path(OUTPUT_DEBUG_TXT, run_tag)

    saved_index_csv_path = safe_write_with_replace(
        index_csv_path,
        lambda output_path: write_index_csv(output_path, index),
    )
    saved_index_txt_path = safe_write_with_replace(
        index_txt_path,
        lambda output_path: write_text(output_path, build_index_text(index)),
    )
    saved_debug_csv_path = safe_write_with_replace(
        debug_csv_path,
        lambda output_path: write_debug_csv(output_path, rows),
    )
    saved_debug_txt_path = safe_write_with_replace(
        debug_txt_path,
        lambda output_path: write_text(output_path, build_debug_text(rows)),
    )
    print_summary(rows, index)

    print("\nSaved files:")
    print(saved_index_csv_path)
    print(saved_index_txt_path)
    print(saved_debug_csv_path)
    print(saved_debug_txt_path)


if __name__ == "__main__":
    main()

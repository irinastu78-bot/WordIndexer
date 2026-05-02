from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
import re
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import normalize_text, read_csv_rows, write_csv, write_text
from scripts.build_draft_author_index_en import (
    DOCX_PATH,
    extract_authors_from_block,
    looks_like_service_title,
    normalize_author_line,
    safe_write_with_replace,
)


OUTPUT_DIR = ROOT_DIR / "output"
SOURCE_PAIRS_CSV = "en_title_author_pairs_debug.csv"
OUTPUT_CSV = "draft_toc_en.csv"
OUTPUT_TXT = "draft_toc_en.txt"
COMPACT_LOCANT_IN_TITLE_RE = re.compile(
    r"(?P<head>\d+)\((?P<inner>\d+)(?P<inside>[A-Za-z][A-Za-z0-9.\[\]]*)\)"
    r"(?P<locants>\d+(?:,\d+)+)(?P<tail>[A-Za-z][A-Za-z0-9.\[\]]*)"
)


@dataclass
class DraftEnTocRow:
    article_no: int | None
    page: int | None
    title: str
    authors_raw: str
    authors_normalized: str
    authors: str
    parsed_authors_count: int
    source_used: str
    source_status: str
    warnings: list[str]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build draft EN TOC artifacts from EN title-author pair debug CSV."
    )
    parser.add_argument(
        "--docx",
        type=Path,
        default=DOCX_PATH,
        help="Accepted for pipeline consistency; this draft step reads existing CSV artifacts only.",
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


def parse_int(value: str) -> int | None:
    normalized = normalize_text(value)
    if normalized.isdigit():
        return int(normalized)
    return None


def normalize_toc_plain_text(value: str) -> str:
    text = value or ""
    text = text.replace("\ufeff", "")
    text = text.replace("\r", "\n").replace("\x0b", "\n").replace("\x0c", "\n")
    text = text.replace("\xa0", " ")
    text = text.replace("\x1e", "-")
    text = text.replace("\x00", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def restore_compact_locant_hyphens_in_title(title: str) -> tuple[str, bool]:
    repaired = False

    def replace_match(match: re.Match[str]) -> str:
        nonlocal repaired
        repaired = True
        return (
            f"{match.group('head')}-"
            f"({match.group('inner')}-{match.group('inside')})-"
            f"{match.group('locants')}-{match.group('tail')}"
        )

    return COMPACT_LOCANT_IN_TITLE_RE.sub(replace_match, title), repaired


def clean_title(value: str) -> tuple[str, bool]:
    title = normalize_toc_plain_text(value)
    return restore_compact_locant_hyphens_in_title(title)


def clean_authors(value: str) -> tuple[str, str, int]:
    raw = normalize_toc_plain_text(value)
    normalized = normalize_author_line(raw)
    parsed_authors = extract_authors_from_block(normalized)
    if parsed_authors:
        return normalized, ", ".join(parsed_authors), len(parsed_authors)
    return normalized, normalized, 0


def build_row(source: dict[str, str]) -> DraftEnTocRow:
    title, repaired_title_hyphens = clean_title(source.get("en_title_candidate", ""))
    authors_raw = normalize_toc_plain_text(source.get("en_author_candidate", ""))
    authors_normalized, authors, parsed_authors_count = clean_authors(authors_raw)
    article_no = parse_int(source.get("article_no", ""))
    page = parse_int(source.get("article_page", ""))
    source_used = normalize_text(source.get("source_used", ""))
    source_status = normalize_text(source.get("status", ""))

    warnings: list[str] = []
    if article_no is None:
        warnings.append("missing_article_no")
    if page is None:
        warnings.append("missing_page")
    if not title:
        warnings.append("missing_title")
    elif looks_like_service_title(title):
        warnings.append("service_title_candidate")
    if repaired_title_hyphens:
        warnings.append("repaired_compact_title_hyphens")
    if not authors_raw:
        warnings.append("missing_authors_raw")
    elif not authors:
        warnings.append("missing_authors_cleaned")
    elif parsed_authors_count == 0:
        warnings.append("authors_parse_failed")
    if source_status != "found_title_and_author":
        warnings.append(f"source_status:{source_status or 'empty'}")

    return DraftEnTocRow(
        article_no=article_no,
        page=page,
        title=title,
        authors_raw=authors_raw,
        authors_normalized=authors_normalized,
        authors=authors,
        parsed_authors_count=parsed_authors_count,
        source_used=source_used,
        source_status=source_status,
        warnings=warnings,
    )


def build_rows(source_csv_path: Path) -> list[DraftEnTocRow]:
    if not source_csv_path.exists():
        raise FileNotFoundError(f"File not found: {source_csv_path}")
    return [build_row(row) for row in read_csv_rows(source_csv_path)]


def write_toc_csv(path: Path, rows: list[DraftEnTocRow]) -> None:
    csv_rows = (
        [
            row.article_no if row.article_no is not None else "",
            row.page if row.page is not None else "",
            row.title,
            row.authors,
            row.authors_raw,
            row.authors_normalized,
            row.parsed_authors_count,
            row.source_used,
            row.source_status,
            " | ".join(row.warnings),
        ]
        for row in rows
    )
    write_csv(
        path,
        [
            "article_no",
            "page",
            "title",
            "authors",
            "authors_raw",
            "authors_normalized",
            "parsed_authors_count",
            "source_used",
            "source_status",
            "warnings",
        ],
        csv_rows,
    )


def build_toc_text(rows: list[DraftEnTocRow]) -> str:
    blocks: list[str] = []
    for row in rows:
        page_text = str(row.page) if row.page is not None else "?"
        title_line = f"{row.title} .... {page_text}" if row.title else f"<missing title> .... {page_text}"
        author_line = row.authors or "<missing authors>"
        blocks.append(f"{title_line}\n{author_line}")
    return "\n\n".join(blocks) + ("\n" if blocks else "")


def print_summary(rows: list[DraftEnTocRow]) -> None:
    missing_title = [row for row in rows if not row.title]
    missing_page = [row for row in rows if row.page is None]
    suspicious = [row for row in rows if row.warnings]

    print("=" * 100)
    print(f"EN TOC ENTRIES:          {len(rows)}")
    print(f"MISSING TITLES:          {len(missing_title)}")
    print(f"MISSING PAGES:           {len(missing_page)}")
    print(f"SUSPICIOUS ROWS:         {len(suspicious)}")
    print("=" * 100)

    if suspicious:
        print("\nSuspicious rows:")
        for row in suspicious[:50]:
            article_text = row.article_no if row.article_no is not None else "?"
            page_text = row.page if row.page is not None else "?"
            warning_text = " | ".join(row.warnings)
            title_text = row.title or "<missing title>"
            print(f"- article {article_text}, page {page_text}: {warning_text}: {title_text}")
        if len(suspicious) > 50:
            print(f"... {len(suspicious) - 50} more")


def main() -> None:
    args = parse_args()
    run_tag = args.run_tag.strip()

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    source_csv_path = resolve_output_path(SOURCE_PAIRS_CSV, run_tag)
    csv_path = resolve_output_path(OUTPUT_CSV, run_tag)
    txt_path = resolve_output_path(OUTPUT_TXT, run_tag)

    print(f"SOURCE PAIRS CSV: {source_csv_path}")
    print(f"RUN TAG:          {run_tag or '(none)'}")

    rows = build_rows(source_csv_path)

    saved_csv_path = safe_write_with_replace(csv_path, lambda output_path: write_toc_csv(output_path, rows))
    saved_txt_path = safe_write_with_replace(txt_path, lambda output_path: write_text(output_path, build_toc_text(rows)))

    print_summary(rows)

    print("\nSaved files:")
    print(saved_csv_path)
    print(saved_txt_path)


if __name__ == "__main__":
    main()

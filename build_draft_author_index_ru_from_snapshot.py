from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
import json
from pathlib import Path
import re
from typing import Optional

from scripts.build_draft_author_index_ru import (
    DraftArticleAuthorRow,
    NAME_EXTRACTION_PATTERNS,
    build_debug_text,
    extract_author_block_from_paragraph,
    extract_authors_from_block,
    reparse_row,
    strip_trailing_author_block,
    write_debug_csv,
)

@dataclass
class SnapshotUpstreamRow:
    article_no: int
    page: Optional[int]
    title_paragraph_index: Optional[int]
    title_paragraph_text: str
    status: str


@dataclass
class SnapshotParagraphRow:
    paragraph_index: int
    text: str


@dataclass
class EnrichmentRow:
    article_no: int
    title_paragraph_index: Optional[int]
    looks_like_title_plus_authors_same_paragraph: bool
    run_fragments: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build snapshot-based RU draft author index debug output."
    )
    parser.add_argument(
        "--run-tag",
        dest="run_tag",
        default="",
        help="Optional prefix for input/output artifact names.",
    )
    return parser.parse_args()


def resolve_output_path(base_name: str, run_tag: str) -> Path:
    clean_run_tag = run_tag.strip()
    file_name = f"{clean_run_tag}_{base_name}" if clean_run_tag else base_name
    return Path("output") / file_name


def clean_text(value: str) -> str:
    return " ".join((value or "").split()).strip()


def parse_int(value: str) -> Optional[int]:
    text = clean_text(value)
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        return None


def parse_bool(value: str) -> bool:
    return clean_text(value).casefold() in {"1", "true", "yes"}


def load_upstream_rows(path: Path) -> list[SnapshotUpstreamRow]:
    rows: list[SnapshotUpstreamRow] = []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for item in reader:
            rows.append(
                SnapshotUpstreamRow(
                    article_no=parse_int(item.get("article_no", "")) or 0,
                    page=parse_int(item.get("page", "")),
                    title_paragraph_index=parse_int(item.get("title_paragraph_index", "")),
                    title_paragraph_text=clean_text(item.get("title_paragraph_text", "")),
                    status=clean_text(item.get("status", "")),
                )
            )
    return rows


def load_snapshot_rows(path: Path) -> list[SnapshotParagraphRow]:
    rows: list[SnapshotParagraphRow] = []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for item in reader:
            paragraph_index = parse_int(item.get("paragraph_index", ""))
            if paragraph_index is None:
                continue
            rows.append(
                SnapshotParagraphRow(
                    paragraph_index=paragraph_index,
                    text=clean_text(item.get("text", "")),
                )
            )
    return rows


def load_enrichment_rows(path: Path) -> list[EnrichmentRow]:
    rows: list[EnrichmentRow] = []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for item in reader:
            rows.append(
                EnrichmentRow(
                    article_no=parse_int(item.get("article_no", "")) or 0,
                    title_paragraph_index=parse_int(item.get("title_paragraph_index", "")),
                    looks_like_title_plus_authors_same_paragraph=parse_bool(
                        item.get("looks_like_title_plus_authors_same_paragraph", "")
                    ),
                    run_fragments=item.get("run_fragments", "") or "",
                )
            )
    return rows


def build_snapshot_index(rows: list[SnapshotParagraphRow]) -> dict[int, SnapshotParagraphRow]:
    return {row.paragraph_index: row for row in rows}


def build_enrichment_index(
    rows: list[EnrichmentRow],
) -> tuple[dict[int, EnrichmentRow], dict[int, EnrichmentRow]]:
    by_article_no = {row.article_no: row for row in rows if row.article_no > 0}
    by_title_paragraph_index = {
        row.title_paragraph_index: row
        for row in rows
        if row.title_paragraph_index is not None
    }
    return by_article_no, by_title_paragraph_index


def is_reliable_author_suffix(suffix: str) -> bool:
    parsed_authors = extract_authors_from_block(suffix)
    if not parsed_authors:
        return False

    reliable_single_author = (
        len(parsed_authors) == 1
        and (
            "." in suffix
            or "," in suffix
            or ";" in suffix
            or "*" in suffix
            or bool(re.search(r"\d", suffix))
        )
    )
    return len(parsed_authors) >= 2 or reliable_single_author


def shift_start_to_reliable_author_match(text: str, start: int) -> int:
    suffix = clean_text(text[start:])
    if not suffix:
        return start

    match_candidates: list[tuple[int, str]] = []
    for _, pattern in NAME_EXTRACTION_PATTERNS:
        for match in pattern.finditer(suffix):
            candidate_suffix = clean_text(suffix[match.start():])
            if not candidate_suffix:
                continue
            if not is_reliable_author_suffix(candidate_suffix):
                continue
            match_candidates.append((match.start(), candidate_suffix))

    if not match_candidates:
        return start

    match_start, _ = min(match_candidates, key=lambda item: item[0])
    if match_start <= 0:
        return start

    leading = clean_text(suffix[:match_start])
    if len(leading.split()) != 1:
        return start
    if any(char in leading for char in ",;:*"):
        return start
    if any(char.isdigit() for char in leading):
        return start

    return start + match_start


def find_author_tail_in_title_paragraph(text: str) -> tuple[str, str]:
    normalized = clean_text(text)
    if not normalized:
        return "", ""

    candidates: list[tuple[int, str, str]] = []
    for _, pattern in NAME_EXTRACTION_PATTERNS:
        for match in pattern.finditer(normalized):
            start = match.start()
            if start < len(normalized) * 0.4:
                continue

            prefix = clean_text(normalized[:start])
            suffix = clean_text(normalized[start:])
            if len(prefix.split()) < 2 or not suffix:
                continue

            if not is_reliable_author_suffix(suffix):
                continue

            candidates.append((start, prefix, suffix))

    if not candidates:
        return normalized, extract_author_block_from_paragraph(normalized)

    start, prefix, suffix = min(candidates, key=lambda item: item[0])
    adjusted_start = shift_start_to_reliable_author_match(normalized, start)
    title_only = clean_text(normalized[:adjusted_start])
    suffix = clean_text(normalized[adjusted_start:])
    if not title_only:
        return normalized, extract_author_block_from_paragraph(normalized)
    return title_only, suffix


def pick_enrichment_row(
    hit: SnapshotUpstreamRow,
    enrichment_by_article_no: dict[int, EnrichmentRow],
    enrichment_by_title_paragraph_index: dict[int, EnrichmentRow],
) -> Optional[EnrichmentRow]:
    if hit.article_no in enrichment_by_article_no:
        return enrichment_by_article_no[hit.article_no]
    if hit.title_paragraph_index is not None:
        return enrichment_by_title_paragraph_index.get(hit.title_paragraph_index)
    return None


def parse_run_fragments_json(raw_value: str) -> list[dict[str, object]]:
    text = clean_text(raw_value)
    if not text:
        return []
    try:
        value = json.loads(text)
    except json.JSONDecodeError:
        return []
    if not isinstance(value, list):
        return []
    fragments: list[dict[str, object]] = []
    for item in value:
        if not isinstance(item, dict):
            continue
        fragments.append(item)
    return fragments


def parse_fragment_font_size(fragment: dict[str, object]) -> Optional[float]:
    value = fragment.get("font_size")
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def join_fragment_texts(fragments: list[dict[str, object]]) -> str:
    parts = [clean_text(str(fragment.get("text", ""))) for fragment in fragments]
    return clean_text(" ".join(part for part in parts if part))


def split_title_and_authors_from_enrichment(
    enrichment_row: Optional[EnrichmentRow],
) -> Optional[tuple[str, str]]:
    if enrichment_row is None:
        return None
    if not enrichment_row.looks_like_title_plus_authors_same_paragraph:
        return None

    raw_fragments = parse_run_fragments_json(enrichment_row.run_fragments)
    fragments = [
        fragment
        for fragment in raw_fragments
        if clean_text(str(fragment.get("text", "")))
    ]
    if len(fragments) < 2:
        return None

    split_index: Optional[int] = None
    for idx in range(1, len(fragments)):
        previous_font_size = parse_fragment_font_size(fragments[idx - 1])
        current_font_size = parse_fragment_font_size(fragments[idx])
        if previous_font_size is None or current_font_size is None:
            continue
        if previous_font_size >= 13.5 and current_font_size <= 12.5:
            split_index = idx
            break

    if split_index is None:
        return None

    article_title = join_fragment_texts(fragments[:split_index])
    raw_author_line = join_fragment_texts(fragments[split_index:])
    if not article_title or not raw_author_line:
        return None
    return article_title, raw_author_line


def find_next_title_index(rows: list[SnapshotUpstreamRow], current_idx: int) -> Optional[int]:
    for next_row in rows[current_idx + 1 :]:
        if next_row.title_paragraph_index is not None:
            return next_row.title_paragraph_index
    return None


def find_next_nonempty_paragraph_text(
    snapshot_rows_by_index: dict[int, SnapshotParagraphRow],
    start_index: int,
    stop_before_index: Optional[int],
) -> str:
    current_index = start_index + 1
    while stop_before_index is None or current_index < stop_before_index:
        row = snapshot_rows_by_index.get(current_index)
        if row is None:
            current_index += 1
            continue
        if row.text:
            return row.text
        current_index += 1
    return ""


def build_rows_from_snapshot(
    upstream_rows: list[SnapshotUpstreamRow],
    snapshot_rows_by_index: dict[int, SnapshotParagraphRow],
    enrichment_by_article_no: dict[int, EnrichmentRow],
    enrichment_by_title_paragraph_index: dict[int, EnrichmentRow],
) -> list[DraftArticleAuthorRow]:
    rows: list[DraftArticleAuthorRow] = []

    for idx, hit in enumerate(upstream_rows):
        next_title_index = find_next_title_index(upstream_rows, idx)
        next_paragraph_text = ""
        if hit.title_paragraph_index is not None:
            next_paragraph_text = find_next_nonempty_paragraph_text(
                snapshot_rows_by_index,
                start_index=hit.title_paragraph_index,
                stop_before_index=next_title_index,
            )

        enrichment_row = pick_enrichment_row(
            hit,
            enrichment_by_article_no,
            enrichment_by_title_paragraph_index,
        )
        split_from_enrichment = split_title_and_authors_from_enrichment(enrichment_row)
        if split_from_enrichment is not None:
            article_title, title_author_block = split_from_enrichment
        else:
            article_title, title_author_block = find_author_tail_in_title_paragraph(hit.title_paragraph_text)
        next_author_block = extract_author_block_from_paragraph(next_paragraph_text)
        if not title_author_block:
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


def print_summary(rows: list[DraftArticleAuthorRow]) -> None:
    title_count = sum(1 for row in rows if row.source_used == "title_paragraph")
    next_count = sum(1 for row in rows if row.source_used == "next_paragraph")
    not_found_count = sum(1 for row in rows if row.source_used == "not_found")

    print("=" * 100)
    print(f"RU ARTICLES PROCESSED:  {len(rows)}")
    print(f"TITLE PARAGRAPH USED:   {title_count}")
    print(f"NEXT PARAGRAPH USED:    {next_count}")
    print(f"NOT FOUND:              {not_found_count}")
    print("=" * 100)


def main() -> int:
    args = parse_args()
    run_tag = args.run_tag.strip()
    upstream_csv_path = resolve_output_path("author_title_paragraph_ru_debug_from_snapshot.csv", run_tag)
    snapshot_csv_path = resolve_output_path("doc_paragraph_snapshot.csv", run_tag)
    enrichment_csv_path = resolve_output_path("ru_title_paragraph_structure_debug.csv", run_tag)
    output_csv_path = resolve_output_path("draft_author_index_ru_debug_from_snapshot.csv", run_tag)
    output_txt_path = resolve_output_path("draft_author_index_ru_debug_from_snapshot.txt", run_tag)

    if not upstream_csv_path.exists():
        raise FileNotFoundError(f"File not found: {upstream_csv_path}")
    if not snapshot_csv_path.exists():
        raise FileNotFoundError(f"File not found: {snapshot_csv_path}")
    if not enrichment_csv_path.exists():
        raise FileNotFoundError(f"File not found: {enrichment_csv_path}")

    upstream_rows = load_upstream_rows(upstream_csv_path)
    snapshot_rows = load_snapshot_rows(snapshot_csv_path)
    enrichment_rows = load_enrichment_rows(enrichment_csv_path)
    snapshot_rows_by_index = build_snapshot_index(snapshot_rows)
    enrichment_by_article_no, enrichment_by_title_paragraph_index = build_enrichment_index(
        enrichment_rows
    )

    rows = build_rows_from_snapshot(
        upstream_rows,
        snapshot_rows_by_index,
        enrichment_by_article_no,
        enrichment_by_title_paragraph_index,
    )

    output_csv_path.parent.mkdir(parents=True, exist_ok=True)
    write_debug_csv(output_csv_path, rows)
    output_txt_path.write_text(build_debug_text(rows), encoding="utf-8")
    print_summary(rows)

    print(f"upstream rows read: {len(upstream_rows)}")
    print(f"snapshot rows read: {len(snapshot_rows)}")
    print(f"enrichment rows read: {len(enrichment_rows)}")
    print(f"csv: {output_csv_path.resolve()}")
    print(f"txt: {output_txt_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

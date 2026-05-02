import argparse
import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


KEYWORDS_PREFIX = "\u043a\u043b\u044e\u0447\u0435\u0432\u044b\u0435 \u0441\u043b\u043e\u0432\u0430:"
TITLE_STYLE_NAME_RU = "\u0441\u0442\u0438\u043b\u044c2"
TITLE_FONT_SIZE = 14.0
TITLE_FONT_TOLERANCE = 0.2
MAX_TITLE_PAGE_GAP = 1


@dataclass
class SnapshotParagraphRow:
    paragraph_index: int
    page: Optional[int]
    style_name: str
    first_char_font_size: Optional[float]
    alignment: Optional[int]
    text: str


@dataclass
class SnapshotAuthorTitleHit:
    article_no: int
    page: Optional[int]
    keyword_paragraph_index: int
    keyword_paragraph_text: str
    title_paragraph_index: Optional[int]
    title_paragraph_text: str
    status: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build RU title paragraph debug layer from the snapshot CSV."
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


def parse_float(value: str) -> Optional[float]:
    text = clean_text(value)
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def load_snapshot_rows(path: Path) -> list[SnapshotParagraphRow]:
    rows: list[SnapshotParagraphRow] = []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for item in reader:
            rows.append(
                SnapshotParagraphRow(
                    paragraph_index=parse_int(item.get("paragraph_index", "")) or 0,
                    page=parse_int(item.get("page", "")),
                    style_name=clean_text(item.get("style_name", "")),
                    first_char_font_size=parse_float(item.get("first_char_font_size", "")),
                    alignment=parse_int(item.get("alignment", "")),
                    text=clean_text(item.get("text", "")),
                )
            )
    return rows


def is_ru_keyword_paragraph(row: SnapshotParagraphRow) -> bool:
    return row.text.casefold().startswith(KEYWORDS_PREFIX)


def is_title_paragraph(row: SnapshotParagraphRow) -> bool:
    if not row.text:
        return False
    if row.first_char_font_size is None:
        return False
    if abs(row.first_char_font_size - TITLE_FONT_SIZE) > TITLE_FONT_TOLERANCE:
        return False

    style_name = row.style_name.casefold()
    return style_name == TITLE_STYLE_NAME_RU or TITLE_STYLE_NAME_RU in style_name


def find_title_paragraph(rows: list[SnapshotParagraphRow], keyword_idx: int) -> Optional[SnapshotParagraphRow]:
    keyword_row = rows[keyword_idx]

    for idx in range(keyword_idx - 1, -1, -1):
        candidate = rows[idx]
        if is_ru_keyword_paragraph(candidate):
            break
        if not candidate.text:
            continue
        if (
            candidate.page is not None
            and keyword_row.page is not None
            and candidate.page < keyword_row.page - MAX_TITLE_PAGE_GAP
        ):
            break
        if is_title_paragraph(candidate):
            return candidate

    return None


def build_hits(rows: list[SnapshotParagraphRow]) -> list[SnapshotAuthorTitleHit]:
    hits: list[SnapshotAuthorTitleHit] = []

    for idx, row in enumerate(rows):
        if not is_ru_keyword_paragraph(row):
            continue

        article_no = len(hits) + 1
        title_row = find_title_paragraph(rows, idx)
        status = "found_title" if title_row is not None else "missing_title"
        page = row.page if row.page is not None else (title_row.page if title_row is not None else None)

        hits.append(
            SnapshotAuthorTitleHit(
                article_no=article_no,
                page=page,
                keyword_paragraph_index=row.paragraph_index,
                keyword_paragraph_text=row.text,
                title_paragraph_index=title_row.paragraph_index if title_row is not None else None,
                title_paragraph_text=title_row.text if title_row is not None else "",
                status=status,
            )
        )

    return hits


def duplicate_pages(hits: list[SnapshotAuthorTitleHit]) -> list[int]:
    counts: dict[int, int] = {}
    for hit in hits:
        if hit.page is None:
            continue
        counts[hit.page] = counts.get(hit.page, 0) + 1
    return sorted(page for page, count in counts.items() if count > 1)


def page_monotonicity_issues(hits: list[SnapshotAuthorTitleHit]) -> list[str]:
    issues: list[str] = []
    prev_page: Optional[int] = None

    for hit in hits:
        if hit.page is None:
            continue
        if prev_page is not None and hit.page < prev_page:
            issues.append(
                f"article {hit.article_no}: page {hit.page} after page {prev_page}"
            )
        prev_page = hit.page

    return issues


def write_debug_csv(path: Path, hits: list[SnapshotAuthorTitleHit]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(
            [
                "article_no",
                "page",
                "title_paragraph_index",
                "title_paragraph_text",
                "status",
            ]
        )
        for hit in hits:
            writer.writerow(
                [
                    hit.article_no,
                    hit.page if hit.page is not None else "",
                    hit.title_paragraph_index if hit.title_paragraph_index is not None else "",
                    hit.title_paragraph_text,
                    hit.status,
                ]
            )


def build_debug_text(hits: list[SnapshotAuthorTitleHit]) -> str:
    missing_title_count = sum(1 for hit in hits if hit.title_paragraph_index is None)
    duplicate_page_values = duplicate_pages(hits)
    monotonicity_issues = page_monotonicity_issues(hits)

    lines = [
        f"total articles found: {len(hits)}",
        f"missing title count: {missing_title_count}",
        "duplicate pages: "
        + (", ".join(str(page) for page in duplicate_page_values) if duplicate_page_values else "none"),
        "page monotonicity issues: "
        + (" | ".join(monotonicity_issues) if monotonicity_issues else "none"),
        "",
    ]

    for hit in hits:
        page_text = str(hit.page) if hit.page is not None else "?"
        title_index_text = str(hit.title_paragraph_index) if hit.title_paragraph_index is not None else "?"
        lines.extend(
            [
                "=" * 100,
                f"ARTICLE {hit.article_no} | page {page_text}",
                f"STATUS: {hit.status}",
                f"KEYWORD PARAGRAPH INDEX: {hit.keyword_paragraph_index}",
                f"KEYWORD PARAGRAPH: {hit.keyword_paragraph_text or '<empty>'}",
                f"TITLE PARAGRAPH INDEX: {title_index_text}",
                "TITLE PARAGRAPH:",
                hit.title_paragraph_text or "<not found>",
                "",
            ]
        )

    return "\n".join(lines).rstrip() + "\n"


def main() -> int:
    args = parse_args()
    run_tag = args.run_tag.strip()
    input_csv_path = resolve_output_path("doc_paragraph_snapshot.csv", run_tag)
    output_csv_path = resolve_output_path("author_title_paragraph_ru_debug_from_snapshot.csv", run_tag)
    output_txt_path = resolve_output_path("author_title_paragraph_ru_debug_from_snapshot.txt", run_tag)

    if not input_csv_path.exists():
        raise FileNotFoundError(f"File not found: {input_csv_path}")

    rows = load_snapshot_rows(input_csv_path)
    hits = build_hits(rows)

    write_debug_csv(output_csv_path, hits)
    output_txt_path.write_text(build_debug_text(hits), encoding="utf-8")

    print(f"snapshot rows read: {len(rows)}")
    print(f"articles found: {len(hits)}")
    print(f"csv: {output_csv_path.resolve()}")
    print(f"txt: {output_txt_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

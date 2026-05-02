import argparse
import csv
import re
from collections import Counter
from pathlib import Path


MIN_TITLE_LENGTH = 8

AUTHOR_NAME_RE = re.compile(
    r"(?:[*\d]\s*)?"
    r"(?:"
    r"[\u0410-\u042f\u0401][\u0430-\u044f\u0451-]+(?:\s+[\u0410-\u042f\u0401]\.\s*[\u0410-\u042f\u0401]\.)"
    r"|"
    r"[\u0410-\u042f\u0401]\.\s*[\u0410-\u042f\u0401]\.\s*[\u0410-\u042f\u0401][\u0430-\u044f\u0451-]+"
    r")"
    r"(?:\s*[*\d])?"
)


def clean_text(value: str) -> str:
    return " ".join((value or "").split()).strip()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a plain-text RU TOC draft from the draft TOC CSV."
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


def parse_int(value: str):
    text = clean_text(value)
    if not text:
        return None
    match = re.search(r"\d+", text)
    if not match:
        return None
    return int(match.group(0))


def is_suspicious_title(title: str) -> list[str]:
    reasons = []
    if not title:
        reasons.append("empty_ru_title_only")
        return reasons
    if len(title) < MIN_TITLE_LENGTH:
        reasons.append("too_short")
    if AUTHOR_NAME_RE.search(title):
        reasons.append("author_like_pattern")
    return reasons


def main() -> int:
    args = parse_args()
    run_tag = args.run_tag.strip()
    input_csv_path = resolve_output_path("draft_toc_ru.csv", run_tag)
    output_txt_path = resolve_output_path("draft_toc_ru.txt", run_tag)
    output_summary_path = resolve_output_path("draft_toc_ru_text_summary.txt", run_tag)

    rows = []
    with input_csv_path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for raw_row in reader:
            ordinal = parse_int(raw_row.get("ordinal", ""))
            title = clean_text(raw_row.get("ru_title_only", ""))
            page = parse_int(raw_row.get("article_page", ""))
            rows.append(
                {
                    "ordinal": ordinal,
                    "ru_title_only": title,
                    "article_page": page,
                }
            )

    rows.sort(key=lambda row: (row["ordinal"] is None, row["ordinal"] or 0))

    output_txt_path.parent.mkdir(parents=True, exist_ok=True)
    draft_lines = []
    suspicious_rows = []
    page_counter = Counter()

    for row in rows:
        title = row["ru_title_only"]
        page = row["article_page"]
        page_text = "" if page is None else str(page)
        draft_lines.append(f"{title} .... {page_text}")

        if page is not None:
            page_counter[page] += 1

        reasons = is_suspicious_title(title)
        if reasons:
            suspicious_rows.append(
                {
                    "ordinal": row["ordinal"],
                    "title": title,
                    "page": page,
                    "reasons": reasons,
                }
            )

    output_txt_path.write_text("\n".join(draft_lines) + "\n", encoding="utf-8")

    duplicate_pages = sorted(page for page, count in page_counter.items() if count > 1)
    summary_lines = [
        f"total rows read: {len(rows)}",
        f"empty ru_title_only count: {sum(1 for row in rows if not row['ru_title_only'])}",
        "duplicate pages: " + (", ".join(str(page) for page in duplicate_pages) if duplicate_pages else "none"),
        f"suspicious rows count: {len(suspicious_rows)}",
    ]
    if suspicious_rows:
        summary_lines.append("suspicious rows:")
        for row in suspicious_rows:
            ordinal_text = "" if row["ordinal"] is None else str(row["ordinal"])
            page_text = "" if row["page"] is None else str(row["page"])
            reasons_text = ", ".join(row["reasons"])
            summary_lines.append(
                f"ordinal={ordinal_text}\tpage={page_text}\treasons={reasons_text}\ttitle={row['title']}"
            )

    output_summary_path.write_text("\n".join(summary_lines) + "\n", encoding="utf-8")

    print(f"rows read: {len(rows)}")
    print(f"draft text: {output_txt_path.resolve()}")
    print(f"summary: {output_summary_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

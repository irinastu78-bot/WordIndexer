from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import normalize_text, write_csv, write_text


OUTPUT_DIR = ROOT_DIR / "output"
DEFAULT_INPUT_DEBUG_CSV = OUTPUT_DIR / "author_title_paragraph_ru_debug.csv"

OUTPUT_CSV = "author_windows_en_debug.csv"
OUTPUT_TXT = "author_windows_en_debug.txt"


@dataclass
class ArticleWindowRow:
    article_no: int
    article_page: int | None
    ru_title_paragraph_text: str
    window_start_paragraph_index: int | None
    window_end_paragraph_index_exclusive: int | None
    status: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build EN author windows debug CSV/TXT from the RU title paragraph layer."
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


def resolve_input_debug_csv(run_tag: str) -> Path:
    clean_run_tag = run_tag.strip()
    if clean_run_tag:
        tagged_snapshot_path = resolve_output_path(
            "author_title_paragraph_ru_debug_from_snapshot.csv",
            clean_run_tag,
        )
        if tagged_snapshot_path.exists():
            return tagged_snapshot_path

        tagged_legacy_path = resolve_output_path("author_title_paragraph_ru_debug.csv", clean_run_tag)
        if tagged_legacy_path.exists():
            return tagged_legacy_path

    return DEFAULT_INPUT_DEBUG_CSV


def parse_int(value: str) -> int | None:
    normalized = normalize_text(value)
    if normalized.isdigit():
        return int(normalized)
    return None


def read_ru_debug_rows(path: Path) -> list[dict[str, str]]:
    first_line = path.read_text(encoding="utf-8-sig").splitlines()[0]
    delimiter = "," if "," in first_line and ";" not in first_line else ";"

    with path.open("r", encoding="utf-8-sig", newline="") as file_obj:
        reader = csv.DictReader(file_obj, delimiter=delimiter)
        return list(reader)


def build_rows_from_ru_debug_csv(path: Path) -> list[ArticleWindowRow]:
    source_rows = read_ru_debug_rows(path)
    parsed_rows: list[ArticleWindowRow] = []

    for item in source_rows:
        parsed_rows.append(
            ArticleWindowRow(
                article_no=parse_int(item.get("article_no", "")) or 0,
                article_page=parse_int(item.get("page", "")),
                ru_title_paragraph_text=normalize_text(item.get("title_paragraph_text", "")),
                window_start_paragraph_index=parse_int(item.get("title_paragraph_index", "")),
                window_end_paragraph_index_exclusive=None,
                status="missing_ru_title",
            )
        )

    for index, row in enumerate(parsed_rows):
        if row.window_start_paragraph_index is None:
            row.status = "missing_ru_title"
            continue

        next_start: int | None = None
        for next_row in parsed_rows[index + 1 :]:
            if next_row.window_start_paragraph_index is not None:
                next_start = next_row.window_start_paragraph_index
                break

        row.window_end_paragraph_index_exclusive = next_start
        row.status = "window_defined"

    return parsed_rows


def write_debug_csv(path: Path, rows: list[ArticleWindowRow]) -> None:
    csv_rows = (
        [
            row.article_no,
            row.article_page if row.article_page is not None else "",
            row.ru_title_paragraph_text,
            row.window_start_paragraph_index if row.window_start_paragraph_index is not None else "",
            row.window_end_paragraph_index_exclusive if row.window_end_paragraph_index_exclusive is not None else "",
            row.status,
        ]
        for row in rows
    )
    write_csv(
        path,
        [
            "article_no",
            "article_page",
            "ru_title_paragraph_text",
            "window_start_paragraph_index",
            "window_end_paragraph_index_exclusive",
            "status",
        ],
        csv_rows,
    )


def build_debug_text(rows: list[ArticleWindowRow]) -> str:
    blocks: list[str] = []

    for row in rows:
        blocks.append(
            "\n".join(
                [
                    "=" * 100,
                    f"ARTICLE {row.article_no} | page {row.article_page if row.article_page is not None else '?'}",
                    f"RU TITLE PARAGRAPH: {row.ru_title_paragraph_text or '<empty>'}",
                    f"WINDOW START: {row.window_start_paragraph_index if row.window_start_paragraph_index is not None else '?'}",
                    (
                        "WINDOW END EXCLUSIVE: "
                        f"{row.window_end_paragraph_index_exclusive if row.window_end_paragraph_index_exclusive is not None else '<document_end>'}"
                    ),
                    f"STATUS: {row.status}",
                ]
            )
        )

    return "\n\n".join(blocks) + ("\n" if blocks else "")


def print_summary(rows: list[ArticleWindowRow]) -> None:
    defined_count = sum(1 for row in rows if row.status == "window_defined")
    missing_count = sum(1 for row in rows if row.status == "missing_ru_title")

    print("=" * 100)
    print(f"RU ARTICLES READ:        {len(rows)}")
    print(f"WINDOWS DEFINED:         {defined_count}")
    print(f"MISSING RU TITLE:        {missing_count}")
    print("=" * 100)


def main() -> None:
    args = parse_args()
    run_tag = args.run_tag.strip()
    input_path = resolve_input_debug_csv(run_tag)
    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    rows = build_rows_from_ru_debug_csv(input_path)

    csv_path = resolve_output_path(OUTPUT_CSV, run_tag)
    txt_path = resolve_output_path(OUTPUT_TXT, run_tag)

    write_debug_csv(csv_path, rows)
    write_text(txt_path, build_debug_text(rows))
    print_summary(rows)

    print("\nФайлы сохранены:")
    print(csv_path)
    print(txt_path)


if __name__ == "__main__":
    main()

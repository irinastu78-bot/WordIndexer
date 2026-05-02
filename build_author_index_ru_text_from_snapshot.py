from __future__ import annotations

import argparse
from pathlib import Path

from scripts.build_draft_author_index_ru import (
    build_author_index,
    build_draft_rows_from_debug_csv,
    build_index_text,
    write_index_csv,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a clean RU author index text from snapshot debug CSV."
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


def main() -> int:
    args = parse_args()
    run_tag = args.run_tag.strip()

    input_csv_path = resolve_output_path("draft_author_index_ru_debug_from_snapshot.csv", run_tag)
    output_csv_path = resolve_output_path("author_index_ru_from_snapshot.csv", run_tag)
    output_txt_path = resolve_output_path("author_index_ru_from_snapshot.txt", run_tag)

    if not input_csv_path.exists():
        raise FileNotFoundError(f"File not found: {input_csv_path}")

    rows = build_draft_rows_from_debug_csv(input_csv_path)
    index = build_author_index(rows)

    output_csv_path.parent.mkdir(parents=True, exist_ok=True)
    write_index_csv(output_csv_path, index)
    output_txt_path.write_text(build_index_text(index), encoding="utf-8")

    print(f"debug rows read: {len(rows)}")
    print(f"authors in index: {len(index)}")
    print(f"csv: {output_csv_path.resolve()}")
    print(f"txt: {output_txt_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

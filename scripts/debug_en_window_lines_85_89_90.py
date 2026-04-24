from __future__ import annotations

import argparse
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import write_text
from scripts.debug_author_block_en import (
    DOCX_PATH,
    OUTPUT_DIR,
    ArticleWindowRow,
    build_window_lines,
    read_article_windows,
    resolve_input_windows_csv,
    win32,
)
from scripts.debug_en_window_lines_first10 import (
    build_debug_text,
    make_debug_row,
    save_with_fallback,
    write_debug_csv,
)


OUTPUT_CSV = "en_window_lines_85_89_90_debug.csv"
OUTPUT_TXT = "en_window_lines_85_89_90_debug.txt"
TARGET_ARTICLE_NOS = (85, 89, 90)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build raw EN window line debug CSV/TXT for article_no 85, 89, 90 only."
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


def select_target_article_windows(windows: list[ArticleWindowRow]) -> list[ArticleWindowRow]:
    target_set = set(TARGET_ARTICLE_NOS)
    return [window for window in windows if window.article_no in target_set]


def build_rows(docx_path: Path, run_tag: str) -> list:
    if win32 is None:
        raise RuntimeError("РќРµ СѓСЃС‚Р°РЅРѕРІР»РµРЅ pywin32. РЈСЃС‚Р°РЅРѕРІРёС‚Рµ: pip install pywin32")

    input_windows_path = resolve_input_windows_csv(run_tag)
    if not input_windows_path.exists():
        raise FileNotFoundError(f"Р¤Р°Р№Р» РЅРµ РЅР°Р№РґРµРЅ: {input_windows_path}")

    windows = read_article_windows(input_windows_path)
    selected_windows = select_target_article_windows(windows)

    print(f"[info] limiting EN raw window debug to article_no in {TARGET_ARTICLE_NOS}")
    print(f"[info] selected article windows: {len(selected_windows)}")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    doc = None
    rows: list = []

    try:
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        paragraph_total = int(doc.Paragraphs.Count)

        for index, window in enumerate(selected_windows, start=1):
            page_text = window.article_page if window.article_page is not None else "?"
            print(f"[progress] article {index}/{len(selected_windows)} (article_no={window.article_no}, page={page_text})")

            if window.window_start_paragraph_index is None:
                rows.append(make_debug_row(window, None, "<missing_window>"))
                continue

            end_exclusive = window.window_end_paragraph_index_exclusive or (paragraph_total + 1)
            lines = build_window_lines(
                doc,
                window.window_start_paragraph_index,
                min(end_exclusive, paragraph_total + 1),
            )

            if not lines:
                rows.append(make_debug_row(window, None, "<empty_window>"))
                continue

            for line in lines:
                rows.append(make_debug_row(window, line, line.text))

        return rows
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


def main() -> None:
    args = parse_args()
    run_tag = args.run_tag.strip()
    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"Р¤Р°Р№Р» РЅРµ РЅР°Р№РґРµРЅ: {docx_path}")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    rows = build_rows(docx_path, run_tag)

    csv_path = resolve_output_path(OUTPUT_CSV, run_tag)
    txt_path = resolve_output_path(OUTPUT_TXT, run_tag)
    debug_text = build_debug_text(rows)

    saved_csv_path = save_with_fallback(csv_path, lambda output_path: write_debug_csv(output_path, rows))
    saved_txt_path = save_with_fallback(txt_path, lambda output_path: write_text(output_path, debug_text))

    print("\nР¤Р°Р№Р»С‹ СЃРѕС…СЂР°РЅРµРЅС‹:")
    print(saved_csv_path)
    print(saved_txt_path)


if __name__ == "__main__":
    main()

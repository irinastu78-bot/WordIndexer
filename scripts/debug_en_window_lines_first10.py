from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
import sys
from typing import Callable

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import normalize_text, write_csv, write_text
from scripts.debug_author_block_en import (
    DOCX_PATH,
    OUTPUT_DIR,
    ArticleWindowRow,
    WindowLine,
    build_window_lines,
    is_keyword_line,
    is_latin_rich,
    looks_like_affiliation_line,
    looks_like_post_title_author_line,
    looks_like_title_line,
    read_article_windows,
    resolve_input_windows_csv,
    win32,
)


OUTPUT_CSV = "en_window_lines_first10_debug.csv"
OUTPUT_TXT = "en_window_lines_first10_debug.txt"
ARTICLE_NO_LIMIT = 10


@dataclass
class EnWindowLineDebugRow:
    article_no: int
    article_page: int | None
    window_start_paragraph_index: int | None
    window_end_paragraph_index_exclusive: int | None
    paragraph_index: int | None
    line_no: int | None
    paragraph_first_char_font_size: float | None
    paragraph_first_char_italic: bool | None
    latin_rich: str
    keyword_line: str
    title_like: str
    post_title_author_like: str
    affiliation_like: str
    text: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build raw EN window line debug CSV/TXT for article_no <= 10 only."
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


def is_retryable_output_write_error(error: OSError) -> bool:
    if isinstance(error, PermissionError):
        return True
    return getattr(error, "winerror", None) == 32


def build_fallback_output_path(path: Path, suffix_no: int) -> Path:
    return path.with_name(f"{path.stem}_{suffix_no}{path.suffix}")


def save_with_fallback(path: Path, writer: Callable[[Path], None]) -> Path:
    try:
        writer(path)
        return path
    except OSError as error:
        if not is_retryable_output_write_error(error):
            raise

    for suffix_no in range(1, 1000):
        fallback_path = build_fallback_output_path(path, suffix_no)
        try:
            writer(fallback_path)
            print(f"[save] primary path unavailable, saved fallback artifact: {fallback_path}")
            return fallback_path
        except OSError as error:
            if not is_retryable_output_write_error(error):
                raise

    raise RuntimeError(f"Could not save artifact with fallback names near: {path}")


def bool_flag(value: bool) -> str:
    return "yes" if value else "no"


def select_first_10_article_windows(windows: list[ArticleWindowRow]) -> list[ArticleWindowRow]:
    selected: list[ArticleWindowRow] = []

    for window in windows:
        if window.article_no <= ARTICLE_NO_LIMIT:
            selected.append(window)
        if window.article_no == ARTICLE_NO_LIMIT:
            break

    return selected


def make_debug_row(
    window: ArticleWindowRow,
    line: WindowLine | None,
    text: str,
) -> EnWindowLineDebugRow:
    normalized_text = normalize_text(text)
    return EnWindowLineDebugRow(
        article_no=window.article_no,
        article_page=window.article_page,
        window_start_paragraph_index=window.window_start_paragraph_index,
        window_end_paragraph_index_exclusive=window.window_end_paragraph_index_exclusive,
        paragraph_index=line.paragraph_index if line is not None else None,
        line_no=line.line_no if line is not None else None,
        paragraph_first_char_font_size=line.paragraph_first_char_font_size if line is not None else None,
        paragraph_first_char_italic=line.paragraph_first_char_italic if line is not None else None,
        latin_rich=bool_flag(is_latin_rich(normalized_text)),
        keyword_line=bool_flag(is_keyword_line(normalized_text)),
        title_like=bool_flag(looks_like_title_line(normalized_text)),
        post_title_author_like=bool_flag(looks_like_post_title_author_line(normalized_text)),
        affiliation_like=bool_flag(looks_like_affiliation_line(normalized_text)),
        text=normalized_text,
    )


def build_rows(docx_path: Path, run_tag: str) -> list[EnWindowLineDebugRow]:
    if win32 is None:
        raise RuntimeError("Не установлен pywin32. Установите: pip install pywin32")

    input_windows_path = resolve_input_windows_csv(run_tag)
    if not input_windows_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_windows_path}")

    windows = read_article_windows(input_windows_path)
    selected_windows = select_first_10_article_windows(windows)

    print(f"[info] limiting EN raw window debug to article_no <= {ARTICLE_NO_LIMIT}")
    print(f"[info] selected article windows: {len(selected_windows)}")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    doc = None
    rows: list[EnWindowLineDebugRow] = []

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

            if window.article_no == ARTICLE_NO_LIMIT:
                break

        return rows
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


def write_debug_csv(path: Path, rows: list[EnWindowLineDebugRow]) -> None:
    csv_rows = (
        [
            row.article_no,
            row.article_page if row.article_page is not None else "",
            row.window_start_paragraph_index if row.window_start_paragraph_index is not None else "",
            row.window_end_paragraph_index_exclusive if row.window_end_paragraph_index_exclusive is not None else "",
            row.paragraph_index if row.paragraph_index is not None else "",
            row.line_no if row.line_no is not None else "",
            row.paragraph_first_char_font_size if row.paragraph_first_char_font_size is not None else "",
            (
                "true"
                if row.paragraph_first_char_italic is True
                else "false"
                if row.paragraph_first_char_italic is False
                else ""
            ),
            row.latin_rich,
            row.keyword_line,
            row.title_like,
            row.post_title_author_like,
            row.affiliation_like,
            row.text,
        ]
        for row in rows
    )
    write_csv(
        path,
        [
            "article_no",
            "article_page",
            "window_start_paragraph_index",
            "window_end_paragraph_index_exclusive",
            "paragraph_index",
            "line_no",
            "paragraph_first_char_font_size",
            "paragraph_first_char_italic",
            "latin_rich",
            "keyword_line",
            "title_like",
            "post_title_author_like",
            "affiliation_like",
            "text",
        ],
        csv_rows,
    )


def build_debug_text(rows: list[EnWindowLineDebugRow]) -> str:
    if not rows:
        return ""

    blocks: list[str] = []
    current_article_no: int | None = None

    for row in rows:
        if row.article_no != current_article_no:
            current_article_no = row.article_no
            blocks.append(
                "\n".join(
                    [
                        "=" * 100,
                        f"ARTICLE {row.article_no} | page {row.article_page if row.article_page is not None else '?'}",
                        f"WINDOW: {row.window_start_paragraph_index if row.window_start_paragraph_index is not None else '?'} -> {row.window_end_paragraph_index_exclusive if row.window_end_paragraph_index_exclusive is not None else '<document_end>'}",
                    ]
                )
            )

        font_text = (
            str(row.paragraph_first_char_font_size)
            if row.paragraph_first_char_font_size is not None
            else "?"
        )
        italic_text = (
            "true"
            if row.paragraph_first_char_italic is True
            else "false"
            if row.paragraph_first_char_italic is False
            else "?"
        )
        paragraph_text = str(row.paragraph_index) if row.paragraph_index is not None else "?"
        line_text = str(row.line_no) if row.line_no is not None else "?"

        blocks.append(
            "\n".join(
                [
                    (
                        f"p={paragraph_text} | line={line_text} | font={font_text} | italic={italic_text} | "
                        f"latin_rich={row.latin_rich} | keyword_line={row.keyword_line} | "
                        f"title_like={row.title_like} | post_title_author_like={row.post_title_author_like} | "
                        f"affiliation_like={row.affiliation_like}"
                    ),
                    row.text or "<empty>",
                ]
            )
        )

    return "\n\n".join(blocks) + "\n"


def main() -> None:
    args = parse_args()
    run_tag = args.run_tag.strip()
    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"Файл не найден: {docx_path}")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    rows = build_rows(docx_path, run_tag)

    csv_path = resolve_output_path(OUTPUT_CSV, run_tag)
    txt_path = resolve_output_path(OUTPUT_TXT, run_tag)
    debug_text = build_debug_text(rows)

    saved_csv_path = save_with_fallback(csv_path, lambda output_path: write_debug_csv(output_path, rows))
    saved_txt_path = save_with_fallback(txt_path, lambda output_path: write_text(output_path, debug_text))

    print("\nФайлы сохранены:")
    print(saved_csv_path)
    print(saved_txt_path)


if __name__ == "__main__":
    main()

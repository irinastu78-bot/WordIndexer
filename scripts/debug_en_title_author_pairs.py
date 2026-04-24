from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import normalize_text, read_csv_rows, write_csv, write_text
from scripts.debug_author_block_en import (
    DOCX_PATH,
    OUTPUT_DIR,
    WindowLine,
    build_window_lines,
    choose_by_forward_scan,
    choose_by_keyword_and_font14,
    choose_from_keyword_window,
    read_article_windows,
    resolve_input_windows_csv,
    win32,
)


OUTPUT_CSV = "en_title_author_pairs_debug.csv"
OUTPUT_TXT = "en_title_author_pairs_debug.txt"


@dataclass
class EnTitleAuthorPairRow:
    article_no: int
    article_page: int | None
    ru_title: str
    ru_authors: str
    keyword_line: str
    en_title_candidate: str
    en_title_paragraph_index: int | None
    en_title_font_size: float | None
    en_author_candidate: str
    en_author_paragraph_index: int | None
    en_author_font_size: float | None
    en_author_italic: bool | None
    source_used: str
    status: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a compact RU/EN title-author pairing debug artifact for EN author block diagnosis."
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
    parser.add_argument(
        "--max-articles",
        dest="max_articles",
        type=int,
        default=0,
        help="Optional limit for the first N article windows only. Defaults to all articles.",
    )
    parser.add_argument(
        "--article-no-from",
        dest="article_no_from",
        type=int,
        default=None,
        help="Optional lower bound for article_no filtering.",
    )
    parser.add_argument(
        "--article-no-to",
        dest="article_no_to",
        type=int,
        default=None,
        help="Optional upper bound for article_no filtering.",
    )
    return parser.parse_args()


def resolve_output_path(base_name: str, run_tag: str) -> Path:
    clean_run_tag = run_tag.strip()
    file_name = f"{clean_run_tag}_{base_name}" if clean_run_tag else base_name
    return OUTPUT_DIR / file_name


def resolve_ru_debug_csv(run_tag: str) -> Path | None:
    candidates: list[Path] = []
    clean_run_tag = run_tag.strip()

    if clean_run_tag:
        candidates.extend(
            [
                resolve_output_path("draft_author_index_ru_debug_from_snapshot.csv", clean_run_tag),
                resolve_output_path("draft_author_index_ru_debug.csv", clean_run_tag),
            ]
        )

    candidates.extend(
        [
            resolve_output_path("draft_author_index_ru_debug_from_snapshot.csv", ""),
            resolve_output_path("draft_author_index_ru_debug.csv", ""),
        ]
    )

    for path in candidates:
        if path.exists():
            return path

    return None


def parse_int(value: str) -> int | None:
    normalized = normalize_text(value)
    if normalized.isdigit():
        return int(normalized)
    return None


def load_ru_authors_by_article_no(path: Path | None) -> dict[int, str]:
    if path is None or not path.exists():
        return {}

    result: dict[int, str] = {}

    for item in read_csv_rows(path):
        article_no = parse_int(item.get("article_no", ""))
        if article_no is None:
            continue

        parsed_authors = normalize_text(item.get("parsed_authors", ""))
        if parsed_authors:
            result[article_no] = parsed_authors

    return result


def choose_candidate_lines(
    lines: list[WindowLine],
) -> tuple[WindowLine | None, WindowLine | None, WindowLine | None, str, str]:
    keyword_line, backward_title_line, backward_author_line = choose_from_keyword_window(lines)
    font14_keyword_line, font14_title_line, font14_author_line = choose_by_keyword_and_font14(lines)

    fallback_source_used = "keyword_backward_scan" if keyword_line is not None else "forward_scan"
    fallback_title_line = backward_title_line
    fallback_author_line = backward_author_line

    if fallback_title_line is None and fallback_author_line is None:
        fallback_title_line, fallback_author_line = choose_by_forward_scan(lines)

    if font14_title_line is not None and font14_author_line is not None:
        return (
            font14_keyword_line or keyword_line,
            font14_title_line,
            font14_author_line,
            "font14",
            "found_title_and_author",
        )

    if fallback_title_line is not None and fallback_author_line is not None:
        return keyword_line, fallback_title_line, fallback_author_line, fallback_source_used, "found_title_and_author"

    if font14_title_line is not None:
        return (
            font14_keyword_line or keyword_line,
            font14_title_line,
            None,
            "font14_title_only",
            "found_title_only",
        )

    if fallback_title_line is not None:
        return keyword_line, fallback_title_line, None, fallback_source_used, "found_title_only"

    return keyword_line or font14_keyword_line, None, None, "not_found", "not_found"


def filter_windows(
    windows: list,
    max_articles: int = 0,
    article_no_from: int | None = None,
    article_no_to: int | None = None,
) -> list:
    if article_no_from is not None and article_no_to is not None and article_no_from > article_no_to:
        raise ValueError("--article-no-from cannot be greater than --article-no-to")

    if article_no_from is not None or article_no_to is not None:
        selected_windows = [
            window
            for window in windows
            if (article_no_from is None or window.article_no >= article_no_from)
            and (article_no_to is None or window.article_no <= article_no_to)
        ]
        range_from_text = article_no_from if article_no_from is not None else "-inf"
        range_to_text = article_no_to if article_no_to is not None else "+inf"
        print(
            f"[info] filtering by article_no range {range_from_text}..{range_to_text} "
            f"-> {len(selected_windows)} article windows"
        )
        return selected_windows

    if max_articles > 0:
        selected_windows = windows[:max_articles]
        print(f"[info] limiting to first {len(selected_windows)} article windows")
        return selected_windows

    return windows


def build_pair_rows(
    doc,
    paragraph_total: int,
    run_tag: str,
    max_articles: int = 0,
    article_no_from: int | None = None,
    article_no_to: int | None = None,
) -> list[EnTitleAuthorPairRow]:
    windows = read_article_windows(resolve_input_windows_csv(run_tag))
    windows = filter_windows(
        windows,
        max_articles=max_articles,
        article_no_from=article_no_from,
        article_no_to=article_no_to,
    )
    ru_authors_by_article_no = load_ru_authors_by_article_no(resolve_ru_debug_csv(run_tag))
    rows: list[EnTitleAuthorPairRow] = []

    for index, window in enumerate(windows, start=1):
        if index == 1 or index % 5 == 0 or index == len(windows):
            page_text = window.article_page if window.article_page is not None else "?"
            print(f"[progress] article {index}/{len(windows)} (article_no={window.article_no}, page={page_text})")

        if window.window_start_paragraph_index is None:
            rows.append(
                EnTitleAuthorPairRow(
                    article_no=window.article_no,
                    article_page=window.article_page,
                    ru_title=window.ru_title_paragraph_text,
                    ru_authors=ru_authors_by_article_no.get(window.article_no, ""),
                    keyword_line="",
                    en_title_candidate="",
                    en_title_paragraph_index=None,
                    en_title_font_size=None,
                    en_author_candidate="",
                    en_author_paragraph_index=None,
                    en_author_font_size=None,
                    en_author_italic=None,
                    source_used="missing_window",
                    status="missing_window",
                )
            )
            continue

        end_exclusive = window.window_end_paragraph_index_exclusive or (paragraph_total + 1)
        lines = build_window_lines(doc, window.window_start_paragraph_index, min(end_exclusive, paragraph_total + 1))
        keyword_line, title_line, author_line, source_used, status = choose_candidate_lines(lines)

        rows.append(
            EnTitleAuthorPairRow(
                article_no=window.article_no,
                article_page=window.article_page,
                ru_title=window.ru_title_paragraph_text,
                ru_authors=ru_authors_by_article_no.get(window.article_no, ""),
                keyword_line=keyword_line.text if keyword_line is not None else "",
                en_title_candidate=title_line.text if title_line is not None else "",
                en_title_paragraph_index=title_line.paragraph_index if title_line is not None else None,
                en_title_font_size=title_line.paragraph_first_char_font_size if title_line is not None else None,
                en_author_candidate=author_line.text if author_line is not None else "",
                en_author_paragraph_index=author_line.paragraph_index if author_line is not None else None,
                en_author_font_size=author_line.paragraph_first_char_font_size if author_line is not None else None,
                en_author_italic=author_line.paragraph_first_char_italic if author_line is not None else None,
                source_used=source_used,
                status=status,
            )
        )

    return rows


def build_rows(
    docx_path: Path,
    run_tag: str,
    max_articles: int = 0,
    article_no_from: int | None = None,
    article_no_to: int | None = None,
) -> list[EnTitleAuthorPairRow]:
    if win32 is None:
        raise RuntimeError("Не установлен pywin32. Установите: pip install pywin32")

    input_windows_path = resolve_input_windows_csv(run_tag)
    if not input_windows_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_windows_path}")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    doc = None

    try:
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        paragraph_total = int(doc.Paragraphs.Count)
        return build_pair_rows(
            doc,
            paragraph_total,
            run_tag,
            max_articles=max_articles,
            article_no_from=article_no_from,
            article_no_to=article_no_to,
        )
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


def is_retryable_output_write_error(error: OSError) -> bool:
    if isinstance(error, PermissionError):
        return True
    return getattr(error, "winerror", None) == 32


def build_fallback_output_path(path: Path, suffix_no: int) -> Path:
    return path.with_name(f"{path.stem}_{suffix_no}{path.suffix}")


def save_with_fallback(path: Path, writer) -> Path:
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


def write_debug_csv(path: Path, rows: list[EnTitleAuthorPairRow]) -> None:
    csv_rows = (
        [
            row.article_no,
            row.article_page if row.article_page is not None else "",
            row.ru_title,
            row.ru_authors,
            row.keyword_line,
            row.en_title_candidate,
            row.en_title_paragraph_index if row.en_title_paragraph_index is not None else "",
            row.en_title_font_size if row.en_title_font_size is not None else "",
            row.en_author_candidate,
            row.en_author_paragraph_index if row.en_author_paragraph_index is not None else "",
            row.en_author_font_size if row.en_author_font_size is not None else "",
            (
                "true"
                if row.en_author_italic is True
                else "false"
                if row.en_author_italic is False
                else ""
            ),
            row.source_used,
            row.status,
        ]
        for row in rows
    )
    write_csv(
        path,
        [
            "article_no",
            "article_page",
            "ru_title",
            "ru_authors",
            "keyword_line",
            "en_title_candidate",
            "en_title_paragraph_index",
            "en_title_font_size",
            "en_author_candidate",
            "en_author_paragraph_index",
            "en_author_font_size",
            "en_author_italic",
            "source_used",
            "status",
        ],
        csv_rows,
    )


def build_debug_text(rows: list[EnTitleAuthorPairRow]) -> str:
    blocks: list[str] = []

    for row in rows:
        page_text = str(row.article_page) if row.article_page is not None else "?"
        title_index_text = str(row.en_title_paragraph_index) if row.en_title_paragraph_index is not None else "?"
        author_index_text = str(row.en_author_paragraph_index) if row.en_author_paragraph_index is not None else "?"
        title_font_text = str(row.en_title_font_size) if row.en_title_font_size is not None else "?"
        author_font_text = str(row.en_author_font_size) if row.en_author_font_size is not None else "?"
        author_italic_text = (
            "true" if row.en_author_italic is True else "false" if row.en_author_italic is False else "?"
        )

        blocks.append(
            "\n".join(
                [
                    "=" * 100,
                    f"ARTICLE {row.article_no} | page {page_text}",
                    f"STATUS: {row.status}",
                    f"SOURCE USED: {row.source_used}",
                    "",
                    "RU TITLE:",
                    row.ru_title or "<empty>",
                    "",
                    "RU AUTHORS:",
                    row.ru_authors or "<empty>",
                    "",
                    "KEYWORD LINE:",
                    row.keyword_line or "<not found>",
                    "",
                    f"EN TITLE PARAGRAPH INDEX: {title_index_text}",
                    f"EN TITLE FONT SIZE: {title_font_text}",
                    "EN TITLE CANDIDATE:",
                    row.en_title_candidate or "<not found>",
                    "",
                    f"EN AUTHOR PARAGRAPH INDEX: {author_index_text}",
                    f"EN AUTHOR FONT SIZE: {author_font_text}",
                    f"EN AUTHOR ITALIC: {author_italic_text}",
                    "EN AUTHOR CANDIDATE:",
                    row.en_author_candidate or "<not found>",
                ]
            )
        )

    return "\n\n".join(blocks) + ("\n" if blocks else "")


def print_summary(rows: list[EnTitleAuthorPairRow]) -> None:
    title_and_author = sum(1 for row in rows if row.status == "found_title_and_author")
    title_only = sum(1 for row in rows if row.status == "found_title_only")
    not_found = sum(1 for row in rows if row.status == "not_found")
    missing_window = sum(1 for row in rows if row.status == "missing_window")

    print("=" * 100)
    print(f"ARTICLES PROCESSED:      {len(rows)}")
    print(f"TITLE AND AUTHOR FOUND:  {title_and_author}")
    print(f"TITLE ONLY FOUND:        {title_only}")
    print(f"NOT FOUND:               {not_found}")
    print(f"MISSING WINDOW:          {missing_window}")
    print("=" * 100)


def main() -> None:
    args = parse_args()
    run_tag = args.run_tag.strip()
    max_articles = max(0, int(args.max_articles))
    article_no_from = args.article_no_from
    article_no_to = args.article_no_to
    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"Файл не найден: {docx_path}")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    rows = build_rows(
        docx_path,
        run_tag,
        max_articles=max_articles,
        article_no_from=article_no_from,
        article_no_to=article_no_to,
    )

    csv_path = resolve_output_path(OUTPUT_CSV, run_tag)
    txt_path = resolve_output_path(OUTPUT_TXT, run_tag)
    debug_text = build_debug_text(rows)

    saved_csv_path = save_with_fallback(csv_path, lambda output_path: write_debug_csv(output_path, rows))
    saved_txt_path = save_with_fallback(txt_path, lambda output_path: write_text(output_path, debug_text))
    print_summary(rows)

    print("\nФайлы сохранены:")
    print(saved_csv_path)
    print(saved_txt_path)


if __name__ == "__main__":
    main()

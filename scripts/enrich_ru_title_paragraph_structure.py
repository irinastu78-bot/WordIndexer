import argparse
import csv
import json
import os
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pythoncom
import win32com.client

WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER = 1
WD_COLLAPSE_START = 1
INTERNAL_BREAK_MARKER = "[[BR]]"
PARAGRAPH_END_CHARS = {"\r", "\x07"}
INTERNAL_BREAK_CHARS = {"\x0b", "\x0c", "\n"}
TITLE_LIKE_STYLE_HINTS = (
    "title",
    "heading",
    "\u0441\u0442\u0438\u043b\u044c2",
    "\u0437\u0430\u0433\u043e\u043b\u043e\u0432",
)
TITLE_FONT_THRESHOLD = 13.5
AUTHOR_FONT_MAX = 12.5
PROGRESS_EVERY = 10


@dataclass
class InputRow:
    article_no: int
    page: Optional[int]
    title_paragraph_index: Optional[int]
    original_title_paragraph_text: str


@dataclass
class EnrichedRow:
    article_no: int
    page: Optional[int]
    title_paragraph_index: Optional[int]
    original_title_paragraph_text: str
    style_name: str
    first_char_font_size: Optional[float]
    has_internal_breaks: bool
    internal_break_count: int
    text_before_first_internal_break: str
    text_after_first_internal_break: str
    run_fragments: str
    first_fragment_font_size: Optional[float]
    second_fragment_font_size: Optional[float]
    looks_like_title_plus_authors_same_paragraph: bool


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Enrich known RU title paragraphs with narrow Word COM structure data."
    )
    parser.add_argument(
        "--docx",
        dest="docx_path",
        required=True,
        type=Path,
        help="Path to the source .docx file.",
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


def clean_text_preserve_inline_spaces(raw_text: str) -> str:
    text = raw_text.replace("\r", " ").replace("\x07", " ").replace("\t", " ")
    return re.sub(r" +", " ", text).strip()


def clean_break_aware_text(raw_text: str) -> str:
    text = raw_text.replace("\t", " ")
    text = re.sub(
        rf"\s*{re.escape(INTERNAL_BREAK_MARKER)}\s*",
        f" {INTERNAL_BREAK_MARKER} ",
        text,
    )
    text = re.sub(r" +", " ", text)
    return text.strip()


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


def build_break_fields(raw_text: str) -> dict[str, object]:
    sanitized = raw_text
    for char in PARAGRAPH_END_CHARS:
        sanitized = sanitized.replace(char, "")

    internal_break_count = sum(sanitized.count(char) for char in INTERNAL_BREAK_CHARS)
    text_with_break_markers = sanitized
    for char in INTERNAL_BREAK_CHARS:
        text_with_break_markers = text_with_break_markers.replace(char, INTERNAL_BREAK_MARKER)

    text_with_break_markers = clean_break_aware_text(text_with_break_markers)
    if INTERNAL_BREAK_MARKER in text_with_break_markers:
        before_break, after_break = text_with_break_markers.split(INTERNAL_BREAK_MARKER, 1)
        text_before_first_internal_break = clean_text(before_break)
        text_after_first_internal_break = clean_text(after_break)
    else:
        text_before_first_internal_break = clean_text(text_with_break_markers)
        text_after_first_internal_break = ""

    return {
        "has_internal_breaks": internal_break_count > 0,
        "internal_break_count": internal_break_count,
        "text_before_first_internal_break": text_before_first_internal_break,
        "text_after_first_internal_break": text_after_first_internal_break,
    }


def safe_style_name(paragraph) -> str:
    try:
        return str(paragraph.Range.Style or "").strip()
    except Exception:
        return ""


def safe_page_number(rng) -> Optional[int]:
    try:
        dup = rng.Duplicate
        dup.Collapse(WD_COLLAPSE_START)
        page = int(dup.Information(WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER))
    except Exception:
        return None
    return page if page > 0 else None


def safe_first_char_font_size(rng) -> Optional[float]:
    try:
        if rng.Characters.Count < 1:
            return None
        size = rng.Characters(1).Font.Size
    except Exception:
        return None
    if size in (None, ""):
        return None
    try:
        return float(size)
    except (TypeError, ValueError):
        return None


def safe_font_size(range_obj) -> Optional[float]:
    try:
        size = range_obj.Font.Size
    except Exception:
        return None
    if size in (None, ""):
        return None
    try:
        return float(size)
    except (TypeError, ValueError):
        return None


def load_input_rows(path: Path) -> list[InputRow]:
    rows: list[InputRow] = []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for item in reader:
            article_no = parse_int(item.get("article_no", ""))
            if article_no is None:
                continue
            rows.append(
                InputRow(
                    article_no=article_no,
                    page=parse_int(item.get("page", "")),
                    title_paragraph_index=parse_int(item.get("title_paragraph_index", "")),
                    original_title_paragraph_text=clean_text(item.get("title_paragraph_text", "")),
                )
            )
    return rows


def collect_run_fragments(rng) -> list[dict[str, object]]:
    fragments: list[dict[str, object]] = []
    current_text_parts: list[str] = []
    current_font_size: Optional[float] = None

    def flush_current_fragment(*, break_after: bool = False) -> None:
        nonlocal current_text_parts
        nonlocal current_font_size

        if not current_text_parts:
            return

        fragments.append(
            {
                "text": clean_text_preserve_inline_spaces("".join(current_text_parts)),
                "font_size": current_font_size,
                "break_after": break_after,
            }
        )
        current_text_parts = []
        current_font_size = None

    try:
        char_total = int(rng.Characters.Count)
    except Exception:
        char_total = 0

    for char_no in range(1, char_total + 1):
        try:
            char_range = rng.Characters(char_no)
            raw_char = str(char_range.Text)
        except Exception:
            continue

        if not raw_char or raw_char in PARAGRAPH_END_CHARS:
            continue

        if raw_char in INTERNAL_BREAK_CHARS:
            flush_current_fragment(break_after=True)
            continue

        char_text = raw_char.replace("\t", " ")
        char_font_size = safe_font_size(char_range)
        if current_text_parts and current_font_size != char_font_size:
            flush_current_fragment()

        current_text_parts.append(char_text)
        current_font_size = char_font_size

    flush_current_fragment()
    return fragments


def first_two_nonempty_fragment_fonts(fragments: list[dict[str, object]]) -> tuple[Optional[float], Optional[float]]:
    nonempty_fragments = [fragment for fragment in fragments if str(fragment.get("text", "")).strip()]
    first_value = nonempty_fragments[0].get("font_size") if len(nonempty_fragments) >= 1 else None
    second_value = nonempty_fragments[1].get("font_size") if len(nonempty_fragments) >= 2 else None
    return (
        float(first_value) if first_value is not None else None,
        float(second_value) if second_value is not None else None,
    )


def is_title_like(style_name: str, first_char_font_size: Optional[float]) -> bool:
    lowered_style = style_name.casefold()
    if any(hint in lowered_style for hint in TITLE_LIKE_STYLE_HINTS):
        return True
    return first_char_font_size is not None and first_char_font_size >= 14.0


def looks_like_title_plus_authors_same_paragraph(
    *,
    style_name: str,
    first_char_font_size: Optional[float],
    has_internal_breaks: bool,
    first_fragment_font_size: Optional[float],
    second_fragment_font_size: Optional[float],
) -> bool:
    if not is_title_like(style_name, first_char_font_size):
        return False
    if not has_internal_breaks:
        return False
    if first_fragment_font_size is None or second_fragment_font_size is None:
        return False
    if first_fragment_font_size < TITLE_FONT_THRESHOLD:
        return False
    return second_fragment_font_size <= AUTHOR_FONT_MAX


def build_empty_enriched_row(row: InputRow) -> EnrichedRow:
    return EnrichedRow(
        article_no=row.article_no,
        page=row.page,
        title_paragraph_index=row.title_paragraph_index,
        original_title_paragraph_text=row.original_title_paragraph_text,
        style_name="",
        first_char_font_size=None,
        has_internal_breaks=False,
        internal_break_count=0,
        text_before_first_internal_break="",
        text_after_first_internal_break="",
        run_fragments="[]",
        first_fragment_font_size=None,
        second_fragment_font_size=None,
        looks_like_title_plus_authors_same_paragraph=False,
    )


def enrich_row(row: InputRow, document) -> EnrichedRow:
    if row.title_paragraph_index is None:
        return build_empty_enriched_row(row)

    try:
        paragraph = document.Paragraphs(row.title_paragraph_index)
    except Exception:
        return build_empty_enriched_row(row)

    rng = paragraph.Range
    raw_text = str(rng.Text)
    break_fields = build_break_fields(raw_text)
    fragments = collect_run_fragments(rng)
    first_fragment_font_size, second_fragment_font_size = first_two_nonempty_fragment_fonts(fragments)
    style_name = safe_style_name(paragraph)
    first_char_font_size = safe_first_char_font_size(rng)
    actual_page = safe_page_number(rng)

    return EnrichedRow(
        article_no=row.article_no,
        page=actual_page if actual_page is not None else row.page,
        title_paragraph_index=row.title_paragraph_index,
        original_title_paragraph_text=row.original_title_paragraph_text,
        style_name=style_name,
        first_char_font_size=first_char_font_size,
        has_internal_breaks=bool(break_fields["has_internal_breaks"]),
        internal_break_count=int(break_fields["internal_break_count"]),
        text_before_first_internal_break=str(break_fields["text_before_first_internal_break"]),
        text_after_first_internal_break=str(break_fields["text_after_first_internal_break"]),
        run_fragments=json.dumps(fragments, ensure_ascii=False),
        first_fragment_font_size=first_fragment_font_size,
        second_fragment_font_size=second_fragment_font_size,
        looks_like_title_plus_authors_same_paragraph=looks_like_title_plus_authors_same_paragraph(
            style_name=style_name,
            first_char_font_size=first_char_font_size,
            has_internal_breaks=bool(break_fields["has_internal_breaks"]),
            first_fragment_font_size=first_fragment_font_size,
            second_fragment_font_size=second_fragment_font_size,
        ),
    )


def ensure_output_writable(paths: list[Path]) -> None:
    checked_dirs = set()
    for path in paths:
        path.parent.mkdir(parents=True, exist_ok=True)

        parent_key = str(path.parent.resolve())
        if parent_key not in checked_dirs:
            checked_dirs.add(parent_key)
            probe = None
            try:
                with tempfile.NamedTemporaryFile(
                    mode="w",
                    encoding="utf-8",
                    dir=path.parent,
                    prefix="write_probe_",
                    suffix=".tmp",
                    delete=False,
                ) as fh:
                    probe = Path(fh.name)
                    fh.write("ok")
            finally:
                if probe and probe.exists():
                    probe.unlink()

        if path.exists():
            with path.open("a", encoding="utf-8"):
                pass


def make_temp_output_path(final_path: Path) -> Path:
    fd, temp_name = tempfile.mkstemp(
        dir=final_path.parent,
        prefix=final_path.stem + ".",
        suffix=final_path.suffix + ".tmp",
    )
    os.close(fd)
    temp_path = Path(temp_name)
    if temp_path.exists():
        temp_path.unlink()
    return temp_path


def write_output_csv(path: Path, rows: list[EnrichedRow]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(
            [
                "article_no",
                "page",
                "title_paragraph_index",
                "original_title_paragraph_text",
                "style_name",
                "first_char_font_size",
                "has_internal_breaks",
                "internal_break_count",
                "text_before_first_internal_break",
                "text_after_first_internal_break",
                "run_fragments",
                "first_fragment_font_size",
                "second_fragment_font_size",
                "looks_like_title_plus_authors_same_paragraph",
            ]
        )
        for row in rows:
            writer.writerow(
                [
                    row.article_no,
                    row.page if row.page is not None else "",
                    row.title_paragraph_index if row.title_paragraph_index is not None else "",
                    row.original_title_paragraph_text,
                    row.style_name,
                    row.first_char_font_size if row.first_char_font_size is not None else "",
                    row.has_internal_breaks,
                    row.internal_break_count,
                    row.text_before_first_internal_break,
                    row.text_after_first_internal_break,
                    row.run_fragments,
                    row.first_fragment_font_size if row.first_fragment_font_size is not None else "",
                    row.second_fragment_font_size if row.second_fragment_font_size is not None else "",
                    row.looks_like_title_plus_authors_same_paragraph,
                ]
            )


def build_output_text(rows: list[EnrichedRow]) -> str:
    rows_with_title_index = sum(1 for row in rows if row.title_paragraph_index is not None)
    rows_with_internal_breaks = sum(1 for row in rows if row.has_internal_breaks)
    flagged_rows = sum(1 for row in rows if row.looks_like_title_plus_authors_same_paragraph)

    lines = [
        f"total rows read: {len(rows)}",
        f"rows with title_paragraph_index: {rows_with_title_index}",
        f"rows with internal breaks: {rows_with_internal_breaks}",
        f"looks_like_title_plus_authors_same_paragraph count: {flagged_rows}",
        "",
    ]

    for row in rows:
        page_text = str(row.page) if row.page is not None else "?"
        title_idx_text = str(row.title_paragraph_index) if row.title_paragraph_index is not None else "?"
        lines.extend(
            [
                "=" * 100,
                f"ARTICLE {row.article_no} | page {page_text}",
                f"TITLE PARAGRAPH INDEX: {title_idx_text}",
                f"STYLE: {row.style_name or '<empty>'}",
                f"FIRST CHAR FONT SIZE: {row.first_char_font_size if row.first_char_font_size is not None else '<none>'}",
                f"HAS INTERNAL BREAKS: {row.has_internal_breaks}",
                f"INTERNAL BREAK COUNT: {row.internal_break_count}",
                f"FIRST FRAGMENT FONT SIZE: {row.first_fragment_font_size if row.first_fragment_font_size is not None else '<none>'}",
                f"SECOND FRAGMENT FONT SIZE: {row.second_fragment_font_size if row.second_fragment_font_size is not None else '<none>'}",
                f"LOOKS LIKE TITLE+AUTHORS SAME PARAGRAPH: {row.looks_like_title_plus_authors_same_paragraph}",
                "ORIGINAL TITLE PARAGRAPH TEXT:",
                row.original_title_paragraph_text or "<empty>",
                "TEXT BEFORE FIRST INTERNAL BREAK:",
                row.text_before_first_internal_break or "<empty>",
                "TEXT AFTER FIRST INTERNAL BREAK:",
                row.text_after_first_internal_break or "<empty>",
                "RUN FRAGMENTS:",
                row.run_fragments,
                "",
            ]
        )

    return "\n".join(lines).rstrip() + "\n"


def main() -> int:
    args = parse_args()
    docx_path = args.docx_path.resolve()
    run_tag = args.run_tag.strip()
    input_csv_path = resolve_output_path("author_title_paragraph_ru_debug_from_snapshot.csv", run_tag)
    output_csv_path = resolve_output_path("ru_title_paragraph_structure_debug.csv", run_tag)
    output_txt_path = resolve_output_path("ru_title_paragraph_structure_debug.txt", run_tag)

    if not input_csv_path.exists():
        raise FileNotFoundError(f"File not found: {input_csv_path}")

    ensure_output_writable([output_csv_path, output_txt_path])
    temp_csv_path = make_temp_output_path(output_csv_path)
    temp_txt_path = make_temp_output_path(output_txt_path)

    input_rows = load_input_rows(input_csv_path)

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(docx_path), ReadOnly=True)

        enriched_rows: list[EnrichedRow] = []
        print(f"input rows: {len(input_rows)}")
        print(f"document: {docx_path}")

        for idx, row in enumerate(input_rows, start=1):
            enriched_rows.append(enrich_row(row, document))
            if idx % PROGRESS_EVERY == 0 or idx == len(input_rows):
                print(f"processed rows: {idx}/{len(input_rows)}")

        write_output_csv(temp_csv_path, enriched_rows)
        temp_txt_path.write_text(build_output_text(enriched_rows), encoding="utf-8")

        os.replace(temp_csv_path, output_csv_path)
        os.replace(temp_txt_path, output_txt_path)

        print(f"csv: {output_csv_path.resolve()}")
        print(f"txt: {output_txt_path.resolve()}")
        return 0
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()

        if temp_csv_path.exists():
            temp_csv_path.unlink()
        if temp_txt_path.exists():
            temp_txt_path.unlink()


if __name__ == "__main__":
    raise SystemExit(main())

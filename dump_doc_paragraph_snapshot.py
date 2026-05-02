import argparse
import csv
import json
import os
import re
import tempfile
from collections import Counter
from pathlib import Path
from typing import Optional

import pythoncom
import win32com.client


WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER = 1
WD_COLLAPSE_START = 1
PROGRESS_EVERY = 1000
INTERNAL_BREAK_MARKER = "[[BR]]"
PARAGRAPH_END_CHARS = {"\r", "\x07"}
INTERNAL_BREAK_CHARS = {"\x0b", "\x0c", "\n"}
TITLE_LIKE_STYLE_HINTS = (
    "title",
    "heading",
    "\u0441\u0442\u0438\u043b\u044c2",
    "\u0437\u0430\u0433\u043e\u043b\u043e\u0432",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Dump a one-pass Word paragraph snapshot via Word COM."
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
        help="Optional prefix for output artifact names.",
    )
    return parser.parse_args()


def resolve_output_path(base_name: str, run_tag: str) -> Path:
    output_dir = Path("output")
    clean_run_tag = run_tag.strip()
    file_name = f"{clean_run_tag}_{base_name}" if clean_run_tag else base_name
    return output_dir / file_name


def clean_paragraph_text(raw_text: str) -> str:
    text = raw_text.replace("\r", " ").replace("\x07", " ")
    return " ".join(text.split()).strip()


def clean_break_aware_text(raw_text: str) -> str:
    text = raw_text.replace("\t", " ")
    text = re.sub(rf"\s*{re.escape(INTERNAL_BREAK_MARKER)}\s*", f" {INTERNAL_BREAK_MARKER} ", text)
    text = re.sub(r" +", " ", text)
    return text.strip()


def build_break_fields(raw_text: str) -> dict[str, object]:
    sanitized = raw_text
    for char in PARAGRAPH_END_CHARS:
        sanitized = sanitized.replace(char, "")

    internal_break_count = sum(sanitized.count(char) for char in INTERNAL_BREAK_CHARS)
    for char in INTERNAL_BREAK_CHARS:
        sanitized = sanitized.replace(char, f" {INTERNAL_BREAK_MARKER} ")

    text_with_break_markers = clean_break_aware_text(sanitized)
    if INTERNAL_BREAK_MARKER in text_with_break_markers:
        before_break, after_break = text_with_break_markers.split(INTERNAL_BREAK_MARKER, 1)
        text_before_first_internal_break = clean_paragraph_text(before_break)
        text_after_first_internal_break = clean_paragraph_text(after_break)
    else:
        text_before_first_internal_break = clean_paragraph_text(text_with_break_markers)
        text_after_first_internal_break = ""

    return {
        "has_internal_breaks": internal_break_count > 0,
        "internal_break_count": internal_break_count,
        "text_with_break_markers": text_with_break_markers,
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


def safe_alignment(paragraph) -> Optional[int]:
    try:
        return int(paragraph.Alignment)
    except Exception:
        return None


def should_use_detailed_structure(
    *,
    raw_text: str,
    style_name: str,
    first_char_font_size: Optional[float],
) -> bool:
    break_fields = build_break_fields(raw_text)
    if break_fields["has_internal_breaks"]:
        return True

    if first_char_font_size is not None and first_char_font_size >= 14.0:
        return True

    lowered_style = style_name.casefold()
    return any(hint in lowered_style for hint in TITLE_LIKE_STYLE_HINTS)


def collect_paragraph_structure_detailed(rng, raw_text: str, first_char_font_size: Optional[float]) -> dict[str, object]:
    fragments: list[dict[str, object]] = []
    current_text_parts: list[str] = []
    current_font_size: Optional[float] = None
    break_fields = build_break_fields(raw_text)

    def flush_current_fragment(*, break_after: bool = False) -> None:
        nonlocal current_text_parts
        nonlocal current_font_size

        if not current_text_parts:
            return

        fragment_text = "".join(current_text_parts)
        fragments.append(
            {
                "text": clean_text_preserve_inline_spaces(fragment_text),
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

    nonempty_fragments = [fragment for fragment in fragments if str(fragment["text"]).strip()]
    font_sizes = [
        float(fragment["font_size"])
        for fragment in nonempty_fragments
        if fragment["font_size"] is not None
    ]

    return {
        "structure_mode": "detailed",
        "has_internal_breaks": break_fields["has_internal_breaks"],
        "internal_break_count": break_fields["internal_break_count"],
        "text_with_break_markers": break_fields["text_with_break_markers"],
        "text_before_first_internal_break": break_fields["text_before_first_internal_break"],
        "text_after_first_internal_break": break_fields["text_after_first_internal_break"],
        "run_fragments": json.dumps(fragments, ensure_ascii=False),
        "min_font_size_in_paragraph": min(font_sizes) if font_sizes else None,
        "max_font_size_in_paragraph": max(font_sizes) if font_sizes else None,
        "first_nonempty_run_font_size": (
            nonempty_fragments[0]["font_size"] if nonempty_fragments else None
        ),
        "last_nonempty_run_font_size": (
            nonempty_fragments[-1]["font_size"] if nonempty_fragments else None
        ),
    }


def collect_paragraph_structure_fast(raw_text: str, first_char_font_size: Optional[float]) -> dict[str, object]:
    break_fields = build_break_fields(raw_text)
    return {
        "structure_mode": "fast",
        "has_internal_breaks": break_fields["has_internal_breaks"],
        "internal_break_count": break_fields["internal_break_count"],
        "text_with_break_markers": break_fields["text_with_break_markers"],
        "text_before_first_internal_break": break_fields["text_before_first_internal_break"],
        "text_after_first_internal_break": break_fields["text_after_first_internal_break"],
        "run_fragments": "[]",
        "min_font_size_in_paragraph": None,
        "max_font_size_in_paragraph": None,
        "first_nonempty_run_font_size": first_char_font_size,
        "last_nonempty_run_font_size": None,
    }


def clean_text_preserve_inline_spaces(raw_text: str) -> str:
    text = raw_text.replace("\r", " ").replace("\x07", " ").replace("\t", " ")
    return re.sub(r" +", " ", text).strip()


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


def write_summary(
    summary_path: Path,
    total_paragraphs: int,
    empty_paragraphs: int,
    pages_detected_count: int,
    min_page: Optional[int],
    max_page: Optional[int],
    style_counter: Counter,
    paragraphs_with_internal_breaks: int,
    fast_path_count: int,
    detailed_path_count: int,
) -> None:
    top_styles = style_counter.most_common(20)
    lines = [
        f"total paragraphs: {total_paragraphs}",
        f"empty paragraphs count: {empty_paragraphs}",
        f"pages detected count: {pages_detected_count}",
        f"min page: {min_page if min_page is not None else 'none'}",
        f"max page: {max_page if max_page is not None else 'none'}",
        f"distinct styles count: {len(style_counter)}",
        f"paragraphs with internal breaks count: {paragraphs_with_internal_breaks}",
        f"fast path paragraphs count: {fast_path_count}",
        f"detailed path paragraphs count: {detailed_path_count}",
        "top 20 styles by frequency:",
    ]
    for style_name, count in top_styles:
        label = style_name if style_name else "<empty>"
        lines.append(f"{count}\t{label}")
    summary_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    docx_path = args.docx_path.resolve()
    run_tag = args.run_tag.strip()

    final_csv_path = resolve_output_path("doc_paragraph_snapshot.csv", run_tag)
    final_summary_path = resolve_output_path("doc_paragraph_snapshot_summary.txt", run_tag)

    ensure_output_writable([final_csv_path, final_summary_path])

    temp_csv_path = make_temp_output_path(final_csv_path)
    temp_summary_path = make_temp_output_path(final_summary_path)

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(docx_path), ReadOnly=True)

        total_paragraphs = int(document.Paragraphs.Count)
        empty_paragraphs = 0
        detected_pages = set()
        style_counter: Counter = Counter()
        min_page = None
        max_page = None
        paragraphs_with_internal_breaks = 0
        fast_path_count = 0
        detailed_path_count = 0

        print(f"document: {docx_path}")
        print(f"total paragraphs: {total_paragraphs}")

        with temp_csv_path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.writer(fh)
            writer.writerow(
                [
                    "paragraph_index",
                    "page",
                    "style_name",
                    "first_char_font_size",
                    "alignment",
                    "text",
                    "structure_mode",
                    "has_internal_breaks",
                    "internal_break_count",
                    "text_with_break_markers",
                    "text_before_first_internal_break",
                    "text_after_first_internal_break",
                    "min_font_size_in_paragraph",
                    "max_font_size_in_paragraph",
                    "first_nonempty_run_font_size",
                    "last_nonempty_run_font_size",
                    "run_fragments",
                ]
            )

            for idx in range(1, total_paragraphs + 1):
                paragraph = document.Paragraphs(idx)
                rng = paragraph.Range
                raw_text = str(rng.Text)
                text = clean_paragraph_text(rng.Text)
                style_name = safe_style_name(paragraph)
                page = safe_page_number(rng)
                first_char_font_size = safe_first_char_font_size(rng)
                alignment = safe_alignment(paragraph)
                if should_use_detailed_structure(
                    raw_text=raw_text,
                    style_name=style_name,
                    first_char_font_size=first_char_font_size,
                ):
                    paragraph_structure = collect_paragraph_structure_detailed(
                        rng,
                        raw_text,
                        first_char_font_size,
                    )
                    detailed_path_count += 1
                else:
                    paragraph_structure = collect_paragraph_structure_fast(
                        raw_text,
                        first_char_font_size,
                    )
                    fast_path_count += 1

                if not text:
                    empty_paragraphs += 1
                if page is not None:
                    detected_pages.add(page)
                    if min_page is None or page < min_page:
                        min_page = page
                    if max_page is None or page > max_page:
                        max_page = page
                style_counter[style_name] += 1
                if paragraph_structure["has_internal_breaks"]:
                    paragraphs_with_internal_breaks += 1

                writer.writerow(
                    [
                        idx,
                        page,
                        style_name,
                        first_char_font_size,
                        alignment,
                        text,
                        paragraph_structure["structure_mode"],
                        paragraph_structure["has_internal_breaks"],
                        paragraph_structure["internal_break_count"],
                        paragraph_structure["text_with_break_markers"],
                        paragraph_structure["text_before_first_internal_break"],
                        paragraph_structure["text_after_first_internal_break"],
                        paragraph_structure["min_font_size_in_paragraph"],
                        paragraph_structure["max_font_size_in_paragraph"],
                        paragraph_structure["first_nonempty_run_font_size"],
                        paragraph_structure["last_nonempty_run_font_size"],
                        paragraph_structure["run_fragments"],
                    ]
                )

                if idx % PROGRESS_EVERY == 0 or idx == total_paragraphs:
                    print(f"processed paragraphs: {idx}/{total_paragraphs}")

        write_summary(
            temp_summary_path,
            total_paragraphs=total_paragraphs,
            empty_paragraphs=empty_paragraphs,
            pages_detected_count=len(detected_pages),
            min_page=min_page,
            max_page=max_page,
            style_counter=style_counter,
            paragraphs_with_internal_breaks=paragraphs_with_internal_breaks,
            fast_path_count=fast_path_count,
            detailed_path_count=detailed_path_count,
        )

        os.replace(temp_csv_path, final_csv_path)
        os.replace(temp_summary_path, final_summary_path)

        print(f"csv: {final_csv_path.resolve()}")
        print(f"summary: {final_summary_path.resolve()}")
        return 0
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()

        if temp_csv_path.exists():
            temp_csv_path.unlink()
        if temp_summary_path.exists():
            temp_summary_path.unlink()


if __name__ == "__main__":
    raise SystemExit(main())

import argparse
import csv
from dataclasses import dataclass
import os
from pathlib import Path
import re
import tempfile
from typing import List, Optional

import pythoncom
import win32com.client


WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER = 1
WD_COLLAPSE_START = 1
WD_ALIGN_PARAGRAPH_CENTER = 1
PROGRESS_EVERY = 1000

KEYWORDS_PREFIX = "\u043a\u043b\u044e\u0447\u0435\u0432\u044b\u0435 \u0441\u043b\u043e\u0432\u0430:"
STYLE_HINT_TITLE_RU = "\u0437\u0430\u0433\u043e\u043b\u043e\u0432"
STYLE_HINT_AUTHOR_RU = "\u0430\u0432\u0442\u043e\u0440"
STYLE_HINT_KEYWORD_RU = "\u043a\u043b\u044e\u0447"
STYLE_HINT_STYLE2_RU = "\u0441\u0442\u0438\u043b\u044c2"
STYLE_HINT_NORMAL_RU = "\u043e\u0431\u044b\u0447\u043d\u044b\u0439"
MAX_HEADER_LOOKBACK = 12
MAX_HEADER_PAGE_GAP = 1
TITLE_TARGET_FONT_SIZE = 14.0
TITLE_FONT_TOLERANCE = 0.6
FONT_SPLIT_TITLE_MIN = 13.5
FONT_SPLIT_TITLE_MAX = 14.5
FONT_SPLIT_AUTHOR_MIN = 11.0
FONT_SPLIT_AUTHOR_MAX = 12.5
FONT_SPLIT_MIN_RUN = 4
RU_SERVICE_WORDS = {
    "\u0438",
    "\u0432",
    "\u043d\u0430",
    "\u0441",
    "\u043f\u043e",
    "\u0434\u043b\u044f",
    "\u043f\u0440\u0438",
    "\u043e",
    "\u043e\u0431",
    "\u043e\u0442",
    "\u0434\u043e",
    "\u0438\u0437",
    "\u043f\u043e\u0434",
    "\u043c\u0435\u0436\u0434\u0443",
    "\u0430",
    "\u043d\u043e",
    "\u0447\u0442\u043e",
    "\u043a\u0430\u043a",
}
AUTHOR_NAME_RE = re.compile(
    r"(?:[*\d]\s*)?"
    r"(?:"
    r"[\u0410-\u042f\u0401][\u0430-\u044f\u0451-]+(?:\s+[\u0410-\u042f\u0401]\.\s*[\u0410-\u042f\u0401]\.)"
    r"|"
    r"[\u0410-\u042f\u0401]\.\s*[\u0410-\u042f\u0401]\.\s*[\u0410-\u042f\u0401][\u0430-\u044f\u0451-]+"
    r")"
    r"(?:\s*[*\d])?"
)


@dataclass
class ParagraphSnapshot:
    index: int
    text: str
    style_name: str
    page: Optional[int]
    first_char_font_size: Optional[float]
    alignment: Optional[int]


@dataclass
class TocDraftRow:
    ordinal: int
    article_page: Optional[int]
    ru_title_text: str
    ru_title_text_raw: str
    ru_title_only: str
    ru_authors_raw: str
    title_paragraph_index: Optional[int]
    paragraph_style_name: str
    first_char_font_size: Optional[float]
    keyword_paragraph_index: int
    split_method: str
    combined_detected: bool


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a draft RU TOC article list directly from a Word document via COM."
    )
    parser.add_argument(
        "docx_path",
        nargs="?",
        type=Path,
        help="Path to the old .docx file.",
    )
    parser.add_argument(
        "--docx",
        dest="docx_path_flag",
        type=Path,
        help="Path to the old .docx file.",
    )
    parser.add_argument(
        "--run-tag",
        dest="run_tag",
        default="",
        help="Optional prefix for output artifact names.",
    )
    return parser.parse_args()


def resolve_docx_path(args: argparse.Namespace) -> Path:
    docx_path = args.docx_path_flag or args.docx_path
    if docx_path is None:
        raise SystemExit("A .docx path must be provided via positional argument or --docx.")
    return docx_path


def resolve_output_path(base_name: str, run_tag: str) -> Path:
    clean_run_tag = run_tag.strip()
    file_name = f"{clean_run_tag}_{base_name}" if clean_run_tag else base_name
    return Path("output") / file_name


def clean_paragraph_text(raw_text: str) -> str:
    text = raw_text.replace("\r", " ").replace("\x07", " ")
    return " ".join(text.split()).strip()


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


def safe_alignment(paragraph) -> Optional[int]:
    try:
        return int(paragraph.Alignment)
    except Exception:
        return None


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


def snapshot_paragraphs(document, total_paragraphs: int) -> List[ParagraphSnapshot]:
    snapshots: List[ParagraphSnapshot] = []
    for idx in range(1, total_paragraphs + 1):
        paragraph = document.Paragraphs(idx)
        text = clean_paragraph_text(paragraph.Range.Text)
        snapshots.append(
            ParagraphSnapshot(
                index=idx,
                text=text,
                style_name=safe_style_name(paragraph),
                page=safe_page_number(paragraph.Range),
                first_char_font_size=safe_first_char_font_size(paragraph.Range),
                alignment=safe_alignment(paragraph),
            )
        )
        if idx % PROGRESS_EVERY == 0 or idx == total_paragraphs:
            print(f"processed paragraphs: {idx}/{total_paragraphs}")
    return snapshots


def median_value(values: List[float]) -> Optional[float]:
    if not values:
        return None
    ordered = sorted(values)
    mid = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[mid]
    return (ordered[mid - 1] + ordered[mid]) / 2.0


def is_keywords_paragraph(snapshot: ParagraphSnapshot) -> bool:
    return snapshot.text.casefold().startswith(KEYWORDS_PREFIX)


def looks_like_title(snapshot: ParagraphSnapshot) -> bool:
    text = snapshot.text
    if not text:
        return False
    lowered = text.casefold()
    if lowered.startswith(KEYWORDS_PREFIX):
        return False
    if ":" in text and len(text.split()) <= 6:
        return False
    if text.isdigit():
        return False
    if len(text) < 8:
        return False
    return True


def body_text_penalty(snapshot: ParagraphSnapshot) -> float:
    text = snapshot.text.strip()
    if not text:
        return 0.0

    words = text.split()
    lowered_words = [word.strip(".,;:!?()[]\"'").casefold() for word in words]
    service_word_count = sum(1 for word in lowered_words if word in RU_SERVICE_WORDS)

    penalty = 0.0
    if len(words) >= 14:
        penalty += 16.0
    elif len(words) >= 10:
        penalty += 8.0

    if text.endswith("."):
        penalty += 12.0

    if service_word_count >= 4:
        penalty += 10.0
    elif service_word_count >= 3:
        penalty += 5.0

    return penalty


def title_score(candidate: ParagraphSnapshot, keyword_snapshot: ParagraphSnapshot) -> float:
    score = 0.0
    distance = keyword_snapshot.index - candidate.index
    score -= distance * 1.5

    if candidate.page is not None and keyword_snapshot.page is not None:
        page_gap = keyword_snapshot.page - candidate.page
        if page_gap < 0:
            return -10_000.0
        score -= page_gap * 8.0

    style_name = candidate.style_name.casefold()
    if "title" in style_name:
        score += 25.0
    if "heading" in style_name:
        score += 18.0
    if STYLE_HINT_TITLE_RU in style_name:
        score += 18.0
    if STYLE_HINT_STYLE2_RU in style_name:
        score += 40.0
    if "author" in style_name or STYLE_HINT_AUTHOR_RU in style_name:
        score -= 10.0
    if "keyword" in style_name or STYLE_HINT_KEYWORD_RU in style_name:
        score -= 20.0
    if (
        style_name == STYLE_HINT_NORMAL_RU
        or "normal" in style_name
        or "body" in style_name
    ):
        score -= 30.0

    if candidate.first_char_font_size is not None:
        score += candidate.first_char_font_size
        if candidate.first_char_font_size >= 12:
            score += 8.0
        if abs(candidate.first_char_font_size - TITLE_TARGET_FONT_SIZE) <= TITLE_FONT_TOLERANCE:
            score += 35.0

    if candidate.alignment == WD_ALIGN_PARAGRAPH_CENTER:
        score += 4.0

    if len(candidate.text.split()) >= 3:
        score += 3.0
    if len(candidate.text.split()) > 12:
        score -= 6.0

    score -= body_text_penalty(candidate)

    return score


def split_title_and_authors_regex(text: str) -> tuple[str, str, str, bool, bool]:
    raw = text.strip()
    if not raw:
        return "", "", "", False, False

    matches = list(AUTHOR_NAME_RE.finditer(raw))
    if not matches:
        return raw, raw, "", False, False

    combined_detected = False
    for match in matches:
        start = match.start()
        if start < len(raw) * 0.45:
            continue

        prefix = raw[:start].rstrip(" ,;:/")
        suffix = raw[start:].strip()
        if len(prefix.split()) < 2:
            continue

        suffix_matches = list(AUTHOR_NAME_RE.finditer(suffix))
        if not suffix_matches or suffix_matches[0].start() != 0:
            continue

        combined_detected = True
        cleaned_suffix = AUTHOR_NAME_RE.sub("", suffix)
        cleaned_suffix = re.sub(r"[\s,;*0-9().-]+", "", cleaned_suffix)
        if cleaned_suffix:
            continue

        return raw, prefix, suffix, True, True

    return raw, raw, "", combined_detected, False


def split_title_and_authors_by_font(paragraph_range) -> Optional[tuple[str, str, str]]:
    raw_parts = []
    meaningful_entries = []
    raw_offset = 0

    for idx in range(1, paragraph_range.Characters.Count + 1):
        char_range = paragraph_range.Characters(idx)
        char_text = str(char_range.Text)
        raw_parts.append(char_text)

        char_size = None
        try:
            size_value = char_range.Font.Size
            if size_value not in (None, ""):
                char_size = float(size_value)
        except Exception:
            char_size = None

        if char_text not in {" ", "\t", "\r", "\n", "\x07"}:
            meaningful_entries.append(
                {
                    "offset": raw_offset,
                    "text": char_text,
                    "font_size": char_size,
                }
            )
        raw_offset += len(char_text)

    if not meaningful_entries:
        return None

    leading_sizes = []
    for entry in meaningful_entries:
        if entry["font_size"] is not None:
            leading_sizes.append(entry["font_size"])
        if len(leading_sizes) >= 12:
            break

    title_font_size = median_value(leading_sizes)
    if title_font_size is None:
        return None
    if not (FONT_SPLIT_TITLE_MIN <= title_font_size <= FONT_SPLIT_TITLE_MAX):
        return None

    total_meaningful = len(meaningful_entries)
    run_start = None
    run_length = 0

    for idx, entry in enumerate(meaningful_entries):
        size = entry["font_size"]
        is_author_sized = (
            size is not None
            and FONT_SPLIT_AUTHOR_MIN <= size <= FONT_SPLIT_AUTHOR_MAX
            and size <= title_font_size - 1.0
        )
        if is_author_sized:
            if run_start is None:
                run_start = idx
            run_length += 1
        else:
            if run_start is not None and run_length >= FONT_SPLIT_MIN_RUN:
                break
            run_start = None
            run_length = 0

    if run_start is None or run_length < FONT_SPLIT_MIN_RUN:
        return None

    if run_start < total_meaningful * 0.35:
        return None

    raw_text = "".join(raw_parts)
    split_offset = meaningful_entries[run_start]["offset"]
    left_text = clean_paragraph_text(raw_text[:split_offset])
    right_text = clean_paragraph_text(raw_text[split_offset:])

    if len(left_text.split()) < 2 or not right_text:
        return None
    if not AUTHOR_NAME_RE.search(right_text):
        return None

    return clean_paragraph_text(raw_text), left_text, right_text


def split_title_and_authors(paragraph_range, fallback_text: str) -> tuple[str, str, str, bool, str]:
    raw_text = clean_paragraph_text(paragraph_range.Text) if paragraph_range is not None else fallback_text.strip()
    if not raw_text:
        return "", "", "", False, ""

    font_split = split_title_and_authors_by_font(paragraph_range) if paragraph_range is not None else None
    if font_split is not None:
        raw, title_only, authors_raw = font_split
        return raw, title_only, authors_raw, True, "font"

    raw, title_only, authors_raw, combined_detected, split_applied = split_title_and_authors_regex(raw_text)
    if split_applied:
        return raw, title_only, authors_raw, True, "regex"
    if combined_detected:
        return raw, raw, "", True, "fallback"
    return raw, raw, "", False, ""


def find_title_for_keyword(
    paragraphs: List[ParagraphSnapshot],
    keyword_idx: int,
) -> Optional[ParagraphSnapshot]:
    keyword_snapshot = paragraphs[keyword_idx]
    best_candidate: Optional[ParagraphSnapshot] = None
    best_score = -10_000.0

    start_idx = max(0, keyword_idx - MAX_HEADER_LOOKBACK)
    for idx in range(keyword_idx - 1, start_idx - 1, -1):
        candidate = paragraphs[idx]
        if not looks_like_title(candidate):
            continue

        if candidate.page is not None and keyword_snapshot.page is not None:
            if keyword_snapshot.page - candidate.page > MAX_HEADER_PAGE_GAP:
                break

        score = title_score(candidate, keyword_snapshot)
        if score > best_score:
            best_score = score
            best_candidate = candidate

    return best_candidate


def build_rows(document, paragraphs: List[ParagraphSnapshot]) -> List[TocDraftRow]:
    rows: List[TocDraftRow] = []
    for keyword_idx, snapshot in enumerate(paragraphs):
        if not is_keywords_paragraph(snapshot):
            continue
        title_snapshot = find_title_for_keyword(paragraphs, keyword_idx)
        title_text = title_snapshot.text if title_snapshot else ""
        title_range = document.Paragraphs(title_snapshot.index).Range if title_snapshot else None
        ru_title_text_raw, ru_title_only, ru_authors_raw, combined_detected, split_method = split_title_and_authors(
            title_range,
            title_text,
        )
        rows.append(
            TocDraftRow(
                ordinal=len(rows) + 1,
                article_page=title_snapshot.page if title_snapshot else snapshot.page,
                ru_title_text=ru_title_text_raw,
                ru_title_text_raw=ru_title_text_raw,
                ru_title_only=ru_title_only,
                ru_authors_raw=ru_authors_raw,
                title_paragraph_index=title_snapshot.index if title_snapshot else None,
                paragraph_style_name=title_snapshot.style_name if title_snapshot else "",
                first_char_font_size=title_snapshot.first_char_font_size if title_snapshot else None,
                keyword_paragraph_index=snapshot.index,
                split_method=split_method,
                combined_detected=combined_detected,
            )
        )
    return rows


def duplicate_pages(rows: List[TocDraftRow]) -> List[int]:
    counts = {}
    for row in rows:
        if row.article_page is None:
            continue
        counts[row.article_page] = counts.get(row.article_page, 0) + 1
    return sorted(page for page, count in counts.items() if count > 1)


def monotonicity_issues(rows: List[TocDraftRow]) -> List[str]:
    issues: List[str] = []
    prev_page: Optional[int] = None
    for row in rows:
        page = row.article_page
        if page is None:
            continue
        if prev_page is not None and page < prev_page:
            issues.append(f"ordinal {row.ordinal}: page {page} after page {prev_page}")
        prev_page = page
    return issues


def write_csv(rows: List[TocDraftRow], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(
            [
                "ordinal",
                "article_page",
                "ru_title_text",
                "ru_title_text_raw",
                "ru_title_only",
                "ru_authors_raw",
                "title_paragraph_index",
                "paragraph_style_name",
                "first_char_font_size",
                "keyword_paragraph_index",
            ]
        )
        for row in rows:
            writer.writerow(
                [
                    row.ordinal,
                    row.article_page,
                    row.ru_title_text,
                    row.ru_title_text_raw,
                    row.ru_title_only,
                    row.ru_authors_raw,
                    row.title_paragraph_index,
                    row.paragraph_style_name,
                    row.first_char_font_size,
                    row.keyword_paragraph_index,
                ]
            )


def write_summary(rows: List[TocDraftRow], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    empty_titles = sum(1 for row in rows if not row.ru_title_text.strip())
    font_split_applied_count = sum(1 for row in rows if row.split_method == "font")
    regex_split_applied_count = sum(1 for row in rows if row.split_method == "regex")
    split_fallback_count = sum(1 for row in rows if row.split_method == "fallback")
    combined_rows_count = sum(1 for row in rows if row.combined_detected)
    duplicate_page_values = duplicate_pages(rows)
    monotonicity = monotonicity_issues(rows)
    lines = [
        f"total articles found: {len(rows)}",
        f"empty titles count: {empty_titles}",
        f"title_plus_authors_combined_rows_count: {combined_rows_count}",
        f"font_split_applied_count: {font_split_applied_count}",
        f"regex_split_applied_count: {regex_split_applied_count}",
        f"split_fallback_count: {split_fallback_count}",
        "duplicate pages: "
        + (", ".join(str(page) for page in duplicate_page_values) if duplicate_page_values else "none"),
        "page monotonicity issues: "
        + (" | ".join(monotonicity) if monotonicity else "none"),
    ]
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    docx_path = resolve_docx_path(args).resolve()
    run_tag = args.run_tag.strip()
    csv_path = resolve_output_path("draft_toc_ru.csv", run_tag)
    summary_path = resolve_output_path("draft_toc_ru_summary.txt", run_tag)

    ensure_output_writable([csv_path, summary_path])

    temp_csv_path = make_temp_output_path(csv_path)
    temp_summary_path = make_temp_output_path(summary_path)

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(docx_path), ReadOnly=True)

        total_paragraphs = int(document.Paragraphs.Count)
        print(f"document: {docx_path}")
        print(f"total paragraphs: {total_paragraphs}")

        paragraphs = snapshot_paragraphs(document, total_paragraphs)
        rows = build_rows(document, paragraphs)
        write_csv(rows, temp_csv_path)
        write_summary(rows, temp_summary_path)

        os.replace(temp_csv_path, csv_path)
        os.replace(temp_summary_path, summary_path)

        print(f"articles found: {len(rows)}")
        print(f"csv: {csv_path.resolve()}")
        print(f"summary: {summary_path.resolve()}")
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

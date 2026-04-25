from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
import re

from scripts.build_draft_author_index_ru import (
    extract_authors_from_block,
    normalize_author_line,
)
from scripts.insert_keyword_indexes_into_word import (
    insert_heading,
    insert_page_break,
    insert_plain_paragraph,
)

try:
    import win32com.client as win32
except ImportError:
    win32 = None

try:
    import pywintypes
except ImportError:
    pywintypes = None


OUTPUT_DIR = Path("output")
INPUT_DIR = Path("input")
DEFAULT_SOURCE_DOCX = Path("input") / "test1.docx"
DEFAULT_OUTPUT_DOCX = OUTPUT_DIR / "final_with_toc_author.docx"
ORIGINAL_TOC_PATH = INPUT_DIR / "original_toc.txt"
MIN_TOC_TITLE_LENGTH = 1
WD_TAB_ALIGNMENT_RIGHT = 2
WD_TAB_LEADER_DOTS = 1
WD_ALIGN_PARAGRAPH_LEFT = 0
PARAGRAPH_END_CHARS = {"\r", "\x07"}
TOC_PROGRESS_STEP = 10
AUTHOR_PROGRESS_STEP = 25
NBSP = "\u00A0"
SPECIAL_TOC_TITLES = [
    "Лунин Валерий Васильевич",
    "II Лунинские чтения 11 февраля 2026 г.",
    "Статьи по материалам пленарных и стендовых докладов II Лунинских чтений 11 февраля 2026 г.",
]
SPECIAL_TOC_FALLBACK_PAGES = [3, 4, 9]
TITLE_SHORT_WORDS = {"и", "в", "с", "на", "по"}
TOC_INTERNAL_SPACE_AFTER = 0
TOC_ARTICLE_SPACE_AFTER = 10
TOC_SPECIAL_SPACE_AFTER = 10
TOC_TITLE_WRAP_MAX_CHARS = 72
TOC_AUTHOR_LINE_MAX_CHARS = 72
TOC_AUTHOR_LAST_LINE_MAX_CHARS = 72
TOC_AUTHOR_WRAP_TARGET = 72
TOC_AUTHOR_LAST_WRAP_TARGET = 72
TOC_AUTHOR_WRAP_HARD_MAX = 72
TOC_AUTHOR_LAST_HARD_MAX = 72
INITIAL_TOKEN_RE = re.compile(r"\b[A-ZА-ЯЁ][a-zа-яё]{0,2}\.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Insert RU TOC and RU author index text blocks into a Word document."
    )
    parser.add_argument(
        "--docx",
        dest="docx_path",
        type=Path,
        default=DEFAULT_SOURCE_DOCX,
        help="Path to the source .docx file.",
    )
    parser.add_argument(
        "--run-tag",
        dest="run_tag",
        default="",
        help="Optional prefix for tagged input artifacts and output .docx.",
    )
    return parser.parse_args()


def resolve_output_path(base_name: str, run_tag: str) -> Path:
    clean_run_tag = run_tag.strip()
    file_name = f"{clean_run_tag}_{base_name}" if clean_run_tag else base_name
    return OUTPUT_DIR / file_name


def resolve_output_docx_path(run_tag: str) -> Path:
    clean_run_tag = run_tag.strip()
    if not clean_run_tag:
        return DEFAULT_OUTPUT_DOCX
    return OUTPUT_DIR / f"{clean_run_tag}_final_with_toc_author.docx"


def resolve_keyword_indexed_docx_path(run_tag: str) -> Path | None:
    clean_run_tag = run_tag.strip()
    if not clean_run_tag:
        return None

    path = OUTPUT_DIR / f"{clean_run_tag}_with_keyword_indexes.docx"
    if path.exists():
        return path
    return None


def resolve_final_insert_source_path(requested_docx_path: Path, run_tag: str) -> Path:
    keyword_indexed_path = resolve_keyword_indexed_docx_path(run_tag)
    if keyword_indexed_path is not None:
        return keyword_indexed_path
    return requested_docx_path


def iter_error_text_parts(value) -> list[str]:
    parts: list[str] = []

    if value is None:
        return parts

    if isinstance(value, str):
        text = value.strip()
        if text:
            parts.append(text)
        return parts

    if isinstance(value, (list, tuple)):
        for item in value:
            parts.extend(iter_error_text_parts(item))
        return parts

    text = str(value).strip()
    if text:
        parts.append(text)
    return parts


def is_open_document_name_conflict(error: Exception) -> bool:
    if pywintypes is None or not isinstance(error, pywintypes.com_error):
        return False

    message_text = " ".join(part.casefold() for part in iter_error_text_parts(error.args))
    conflict_markers = (
        "нельзя присвоить документу имя уже открытого документа",
        "already open document",
        "name of an open document",
        "cannot assign a document the name of an open document",
    )
    return any(marker in message_text for marker in conflict_markers)


def build_numbered_docx_path(path: Path, index: int) -> Path:
    return path.with_name(f"{path.stem}_{index}{path.suffix}")


def save_document_with_fallback(document, output_docx_path: Path) -> Path:
    target_path = output_docx_path.resolve()

    try:
        document.SaveAs2(str(target_path))
        return target_path
    except Exception as error:
        if not is_open_document_name_conflict(error):
            raise

    fallback_index = 1
    while True:
        fallback_path = build_numbered_docx_path(target_path, fallback_index)
        try:
            document.SaveAs2(str(fallback_path))
            print(f"fallback output docx: {fallback_path}")
            return fallback_path
        except Exception as error:
            if not is_open_document_name_conflict(error):
                raise
            fallback_index += 1


def clean_text(value: str) -> str:
    return " ".join((value or "").split()).strip()


def normalize_toc_title_text(value: str) -> str:
    return (value or "").replace("\x1e", "-")


def parse_int(value: str) -> int | None:
    text = clean_text(value)
    if not text:
        return None
    match = re.search(r"\d+", text)
    if not match:
        return None
    return int(match.group(0))


def parse_bool(value: str) -> bool:
    return clean_text(value).lower() in {"1", "true", "yes"}


def read_nonempty_lines(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    lines = []
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = " ".join(raw_line.split()).strip()
        if line:
            lines.append(line)
    return lines


def read_text_with_fallbacks(path: Path) -> str:
    for encoding in ("utf-8-sig", "cp1251", "utf-16"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def clean_toc_authors(value: str) -> str:
    text = clean_text(value)
    if not text:
        return ""

    normalized = normalize_author_line(text)
    parsed_authors = extract_authors_from_block(normalized)
    if parsed_authors:
        return ", ".join(parsed_authors)

    normalized = re.sub(r"(?<=[A-Za-zА-Яа-яЁё])\d+(?=$|[\s,;])", "", normalized)
    normalized = re.sub(r"\s+,", ",", normalized)
    normalized = re.sub(r",\s*,+", ", ", normalized)
    normalized = re.sub(r",\s*(?=,|$)", "", normalized)
    normalized = re.sub(r"\s*;\s*", "; ", normalized)
    normalized = re.sub(r"\s*,\s*", ", ", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip(" ,;")


def apply_nbsp_to_plain_text(text: str) -> str:
    placeholder = "\uFFF0"
    result = " ".join((text or "").replace(NBSP, placeholder).split()).strip()
    result = result.replace(placeholder, NBSP)
    if not result:
        return ""

    pattern = re.compile(rf"\b({'|'.join(TITLE_SHORT_WORDS)})\s+", flags=re.IGNORECASE)
    return pattern.sub(lambda match: f"{match.group(1)}{NBSP}", result)


def parse_run_fragments(value: str) -> list[dict[str, object]]:
    text = (value or "").strip()
    if not text:
        return []
    try:
        parsed = json.loads(text)
    except Exception:
        return []
    return parsed if isinstance(parsed, list) else []


def load_special_toc_entries() -> list[dict[str, object]]:
    pages = list(SPECIAL_TOC_FALLBACK_PAGES)

    if ORIGINAL_TOC_PATH.exists():
        raw_text = read_text_with_fallbacks(ORIGINAL_TOC_PATH)
        nonempty_lines = [clean_text(line) for line in raw_text.splitlines() if clean_text(line)]
        extracted_pages: list[int] = []

        for line in nonempty_lines:
            normalized = clean_text(line)
            if normalized.casefold() == "оглавление":
                continue

            match = re.search(r"(\d+)\s*$", normalized)
            if match:
                extracted_pages.append(int(match.group(1)))
            if len(extracted_pages) == len(SPECIAL_TOC_TITLES):
                break

        if len(extracted_pages) == len(SPECIAL_TOC_TITLES):
            pages = extracted_pages

    return [
        {
            "title": title,
            "authors": "",
            "page": page,
            "title_paragraph_index": None,
            "is_special": True,
        }
        for title, page in zip(SPECIAL_TOC_TITLES, pages)
    ]


def build_toc_structure_overrides(run_tag: str) -> dict[int, dict[str, str]]:
    path = resolve_output_path("ru_title_paragraph_structure_debug.csv", run_tag)
    if not path.exists():
        return {}

    overrides: dict[int, dict[str, str]] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            paragraph_index = parse_int(row.get("title_paragraph_index", ""))
            if paragraph_index is None:
                continue

            if not parse_bool(row.get("looks_like_title_plus_authors_same_paragraph", "")):
                continue

            fragments = []
            for fragment in parse_run_fragments(row.get("run_fragments", "")):
                text = clean_text(normalize_toc_title_text(str(fragment.get("text", ""))))
                font_size = fragment.get("font_size")
                try:
                    parsed_font_size = float(font_size)
                except (TypeError, ValueError):
                    parsed_font_size = None
                if not text:
                    continue
                fragments.append(
                    {
                        "text": text,
                        "font_size": parsed_font_size,
                    }
                )

            split_index = None
            saw_title_fragment = False
            for idx, fragment in enumerate(fragments):
                font_size = fragment["font_size"]
                if font_size is not None and font_size >= 13.5:
                    saw_title_fragment = True
                    continue
                if saw_title_fragment and font_size is not None and font_size <= 12.5:
                    split_index = idx
                    break

            if split_index is None:
                continue

            title_text = clean_text(" ".join(fragment["text"] for fragment in fragments[:split_index]))
            authors_text = clean_text(" ".join(fragment["text"] for fragment in fragments[split_index:]))
            if not title_text or not authors_text:
                continue

            overrides[paragraph_index] = {
                "title": title_text,
                "authors": authors_text,
            }

    return overrides


def load_toc_entries(run_tag: str) -> list[dict[str, object]]:
    csv_path = resolve_output_path("draft_toc_ru.csv", run_tag)
    if csv_path.exists():
        structure_overrides = build_toc_structure_overrides(run_tag)
        entries: list[dict[str, object]] = load_special_toc_entries()
        with csv_path.open("r", encoding="utf-8-sig", newline="") as fh:
            reader = csv.DictReader(fh)
            for row in reader:
                title_paragraph_index = parse_int(row.get("title_paragraph_index", ""))
                override = (
                    structure_overrides.get(title_paragraph_index)
                    if title_paragraph_index is not None
                    else None
                )

                title = clean_text(
                    normalize_toc_title_text(override["title"] if override is not None else row.get("ru_title_only", ""))
                )
                if len(title) < MIN_TOC_TITLE_LENGTH:
                    continue
                entries.append(
                    {
                        "title": title,
                        "authors": clean_toc_authors(
                            override["authors"] if override is not None else row.get("ru_authors_raw", "")
                        ),
                        "page": parse_int(row.get("article_page", "")),
                        "title_paragraph_index": title_paragraph_index,
                        "is_special": False,
                    }
                )
        return entries

    txt_path = resolve_output_path("draft_toc_ru.txt", run_tag)
    return load_special_toc_entries() + [
        {
            "title": line,
            "authors": "",
            "page": None,
            "title_paragraph_index": None,
            "is_special": False,
        }
        for line in read_nonempty_lines(txt_path)
    ]


def insert_text_block(doc, heading: str, lines: list[str]) -> None:
    insert_heading(doc, heading, level=1)
    for line in lines:
        insert_plain_paragraph(doc, line)


def get_right_tab_stop_position(doc) -> float:
    page_setup = doc.PageSetup
    return float(page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin)


def find_title_source_range(doc, title_paragraph_index: int | None, title_text: str):
    if title_paragraph_index is None or not title_text:
        return None

    try:
        paragraph_range = doc.Paragraphs(title_paragraph_index).Range.Duplicate
    except Exception:
        return None

    target_title = clean_text(title_text)
    if not target_title:
        return None

    raw_offset = 0
    visible_parts: list[str] = []
    matched_end_offset: int | None = None

    try:
        char_total = int(paragraph_range.Characters.Count)
    except Exception:
        return None

    for idx in range(1, char_total + 1):
        try:
            char_range = paragraph_range.Characters(idx)
            raw_char = str(char_range.Text)
        except Exception:
            continue

        char_len = len(raw_char)
        if raw_char not in PARAGRAPH_END_CHARS:
            visible_parts.append(" " if raw_char.isspace() else raw_char)
            if clean_text("".join(visible_parts)) == target_title:
                matched_end_offset = raw_offset + char_len
                break

        raw_offset += char_len

    if matched_end_offset is None or matched_end_offset <= 0:
        return None

    title_range = paragraph_range.Duplicate
    title_range.End = paragraph_range.Start + matched_end_offset
    return title_range


def build_title_char_specs(source_range) -> list[dict[str, object]]:
    specs: list[dict[str, object]] = []

    try:
        char_total = int(source_range.Characters.Count)
    except Exception:
        return specs

    for idx in range(1, char_total + 1):
        try:
            char_range = source_range.Characters(idx)
            raw_char = str(char_range.Text)
        except Exception:
            continue

        if raw_char in PARAGRAPH_END_CHARS:
            continue

        normalized_char = " " if raw_char.isspace() else raw_char
        if normalized_char == " " and (not specs or specs[-1]["text"] == " "):
            continue

        specs.append(
            {
                "text": normalized_char,
                "superscript": bool(char_range.Font.Superscript) if normalized_char != " " else False,
                "subscript": bool(char_range.Font.Subscript) if normalized_char != " " else False,
            }
        )

    while specs and specs[0]["text"] == " ":
        specs.pop(0)
    while specs and specs[-1]["text"] == " ":
        specs.pop()

    return specs


def apply_nbsp_to_title_specs(title_specs: list[dict[str, object]]) -> list[dict[str, object]]:
    specs = [dict(spec) for spec in title_specs]

    for idx, spec in enumerate(specs):
        if spec["text"] != " ":
            continue

        prev_chars: list[str] = []
        scan_idx = idx - 1
        while scan_idx >= 0 and specs[scan_idx]["text"] not in {" ", NBSP}:
            prev_chars.append(str(specs[scan_idx]["text"]))
            scan_idx -= 1

        if not prev_chars:
            continue

        prev_word = "".join(reversed(prev_chars)).casefold()
        if prev_word not in TITLE_SHORT_WORDS:
            continue

        next_idx = idx + 1
        while next_idx < len(specs) and specs[next_idx]["text"] in {" ", NBSP}:
            next_idx += 1
        if next_idx >= len(specs):
            continue

        spec["text"] = NBSP

    return specs


def estimate_title_visual_width(title_specs: list[dict[str, object]]) -> int:
    return len("".join(str(spec["text"]) for spec in title_specs))


def split_title_specs_into_lines(title_specs: list[dict[str, object]]) -> list[list[dict[str, object]]]:
    tokens: list[list[dict[str, object]]] = []
    current_token: list[dict[str, object]] = []

    for spec in title_specs:
        if spec["text"] == " ":
            if current_token:
                tokens.append(current_token)
                current_token = []
            continue
        current_token.append(spec)

    if current_token:
        tokens.append(current_token)

    lines: list[list[dict[str, object]]] = []
    current_line: list[dict[str, object]] = []
    space_spec = {"text": " ", "superscript": False, "subscript": False}

    for token in tokens:
        candidate_line = [*current_line, dict(space_spec), *token] if current_line else list(token)
        if current_line and estimate_title_visual_width(candidate_line) > TOC_TITLE_WRAP_MAX_CHARS:
            lines.append(current_line)
            current_line = list(token)
        else:
            current_line = candidate_line

    if current_line:
        lines.append(current_line)

    return lines


def build_plain_title_specs(text: str) -> list[dict[str, object]]:
    normalized = clean_text(text)
    return [
        {
            "text": char,
            "superscript": False,
            "subscript": False,
        }
        for char in normalized
    ]


def normalize_inserted_title_paragraph(
    paragraph,
    *,
    keep_with_next: bool,
    space_after: int,
) -> None:
    paragraph_range = paragraph.Range.Duplicate
    if paragraph_range.End > paragraph_range.Start:
        paragraph_range.End -= 1

    paragraph_format = paragraph.Range.ParagraphFormat
    paragraph_format.Alignment = WD_ALIGN_PARAGRAPH_LEFT
    paragraph_format.LeftIndent = 0
    paragraph_format.RightIndent = 0
    paragraph_format.FirstLineIndent = 0
    paragraph_format.SpaceBefore = 0
    paragraph_format.SpaceAfter = space_after
    paragraph_format.KeepTogether = True
    paragraph_format.KeepWithNext = keep_with_next
    paragraph_format.WidowControl = False

    try:
        char_total = int(paragraph_range.Characters.Count)
    except Exception:
        return

    for idx in range(1, char_total + 1):
        try:
            char_range = paragraph_range.Characters(idx)
            raw_char = str(char_range.Text)
        except Exception:
            continue

        if raw_char in PARAGRAPH_END_CHARS:
            continue

        preserve_superscript = bool(char_range.Font.Superscript)
        preserve_subscript = bool(char_range.Font.Subscript)

        char_range.Font.Name = "Times New Roman"
        char_range.Font.Size = 12
        char_range.Font.Bold = False
        char_range.Font.Italic = False
        char_range.Font.Superscript = preserve_superscript
        char_range.Font.Subscript = preserve_subscript


def insert_normalized_title_text(
    doc,
    title_specs: list[dict[str, object]],
    *,
    keep_with_next: bool,
    space_after: int,
):
    title_specs = apply_nbsp_to_title_specs(title_specs)
    title_lines = split_title_specs_into_lines(title_specs)
    if not title_lines:
        return None

    inserted_paragraph = None
    for line_index, line_specs in enumerate(title_lines):
        is_last_line = line_index == len(title_lines) - 1
        plain_text = "".join(str(spec["text"]) for spec in line_specs)
        if not plain_text:
            continue

        insert_at = doc.Content.End - 1
        paragraph_range = doc.Range(insert_at, insert_at)
        paragraph_range.InsertAfter(plain_text)
        doc.Range(insert_at + len(plain_text), insert_at + len(plain_text)).InsertParagraphAfter()

        for offset, spec in enumerate(line_specs):
            char_range = doc.Range(insert_at + offset, insert_at + offset + 1)
            char_range.Font.Name = "Times New Roman"
            char_range.Font.Size = 12
            char_range.Font.Bold = False
            char_range.Font.Italic = False
            char_range.Font.Superscript = bool(spec["superscript"])
            char_range.Font.Subscript = bool(spec["subscript"])

        inserted_paragraph = doc.Range(insert_at, insert_at + len(plain_text)).Paragraphs(1)
        normalize_inserted_title_paragraph(
            inserted_paragraph,
            keep_with_next=keep_with_next if is_last_line else True,
            space_after=space_after if is_last_line else TOC_INTERNAL_SPACE_AFTER,
        )

    return inserted_paragraph


def insert_title_with_formatting(
    doc,
    *,
    title: str,
    title_paragraph_index: int | None,
    keep_with_next: bool,
    space_after: int,
):
    source_range = find_title_source_range(doc, title_paragraph_index, title)
    if source_range is None:
        fallback_title = clean_text(title)
        if fallback_title:
            return insert_normalized_title_text(
                doc,
                build_plain_title_specs(fallback_title),
                keep_with_next=keep_with_next,
                space_after=space_after,
            )
        return None

    title_specs = build_title_char_specs(source_range)
    if not title_specs:
        fallback_title = clean_text(title)
        if fallback_title:
            return insert_normalized_title_text(
                doc,
                build_plain_title_specs(fallback_title),
                keep_with_next=keep_with_next,
                space_after=space_after,
            )
        return None

    return insert_normalized_title_text(
        doc,
        title_specs,
        keep_with_next=keep_with_next,
        space_after=space_after,
    )


def split_author_names(authors: str) -> list[str]:
    return [clean_text(item) for item in re.split(r"\s*,\s*", authors) if clean_text(item)]


def format_author_display_name(name: str) -> str:
    normalized = clean_text(name)
    if not normalized:
        return ""
    return re.sub(r"\s+", NBSP, normalized)


def join_author_names(names: list[str]) -> str:
    formatted_names = [formatted for formatted in (format_author_display_name(name) for name in names) if formatted]
    return ", ".join(formatted_names)


def estimate_author_visual_width(text: str) -> float:
    normalized = clean_text(text)
    if not normalized:
        return 0.0

    width = float(len(normalized))
    width -= 0.15 * normalized.count(" ")
    width -= 0.20 * normalized.count(",")
    width -= 0.15 * normalized.count(".")
    width -= 0.10 * normalized.count("-")
    width -= 0.25 * len(INITIAL_TOKEN_RE.findall(normalized))
    return max(width, 0.0)


def score_author_chunk(
    chunk: list[str],
    *,
    is_last: bool,
    remaining_names_count: int,
) -> float:
    line_text = join_author_names(chunk)
    visual_width = estimate_author_visual_width(line_text)

    soft_limit = float(TOC_AUTHOR_LAST_LINE_MAX_CHARS - 2 if is_last else TOC_AUTHOR_LINE_MAX_CHARS - 4)
    hard_limit = float(TOC_AUTHOR_LAST_LINE_MAX_CHARS + 1 if is_last else TOC_AUTHOR_LINE_MAX_CHARS + 1)

    if visual_width <= soft_limit:
        score = (soft_limit - visual_width) ** 2
    elif visual_width <= hard_limit:
        score = (visual_width - soft_limit) ** 2 * 3.5
    else:
        score = 1500.0 + (visual_width - hard_limit) ** 2 * 160.0

    if len(chunk) == 1:
        score += 700.0 if remaining_names_count > 0 else 140.0

    if is_last and remaining_names_count == 0 and visual_width < soft_limit * 0.58:
        score += (soft_limit * 0.58 - visual_width) ** 2 * 2.0

    score += 28.0 if not is_last else 8.0

    return score


def fits_author_line(
    names: list[str],
    *,
    is_last: bool,
) -> bool:
    if not names:
        return True

    visual_width = estimate_author_visual_width(join_author_names(names))
    hard_limit = float(TOC_AUTHOR_LAST_HARD_MAX if is_last else TOC_AUTHOR_WRAP_HARD_MAX)
    return visual_width <= hard_limit


def wrap_author_lines(authors: str) -> list[str]:
    names = split_author_names(authors)
    if not names:
        return []

    line_chunks: list[list[str]] = []
    current_line: list[str] = []

    for name in names:
        candidate_line = [*current_line, name]
        candidate_width = estimate_author_visual_width(join_author_names(candidate_line))

        if not current_line or candidate_width <= float(TOC_AUTHOR_WRAP_HARD_MAX):
            current_line = candidate_line
            continue

        line_chunks.append(current_line)
        current_line = [name]

    if current_line:
        line_chunks.append(current_line)

    lines: list[str] = []
    for index, chunk in enumerate(line_chunks):
        line_text = join_author_names(chunk)
        if index < len(line_chunks) - 1:
            line_text += ","
        lines.append(line_text)

    return lines


def normalize_toc_text_paragraph(
    paragraph,
    *,
    italic: bool,
    keep_with_next: bool,
    space_after: int,
    add_page_tab_stop: bool,
    doc,
) -> None:
    paragraph_range = paragraph.Range.Duplicate
    if paragraph_range.End > paragraph_range.Start:
        paragraph_range.End -= 1

    paragraph_range.Font.Name = "Times New Roman"
    paragraph_range.Font.Size = 12
    paragraph_range.Font.Bold = False
    paragraph_range.Font.Italic = italic

    paragraph_format = paragraph.Range.ParagraphFormat
    paragraph_format.Alignment = WD_ALIGN_PARAGRAPH_LEFT
    paragraph_format.LeftIndent = 0
    paragraph_format.RightIndent = 0
    paragraph_format.FirstLineIndent = 0
    paragraph_format.SpaceBefore = 0
    paragraph_format.SpaceAfter = space_after
    paragraph_format.KeepTogether = True
    paragraph_format.KeepWithNext = keep_with_next
    paragraph_format.WidowControl = False
    paragraph_format.TabStops.ClearAll()

    if add_page_tab_stop:
        paragraph_format.TabStops.Add(
            Position=get_right_tab_stop_position(doc),
            Alignment=WD_TAB_ALIGNMENT_RIGHT,
            Leader=WD_TAB_LEADER_DOTS,
        )


def insert_toc_text_paragraph(
    doc,
    *,
    text: str,
    italic: bool,
    keep_with_next: bool,
    space_after: int,
    page_text: str = "",
):
    visible_text = apply_nbsp_to_plain_text(text)
    line_text = f"{visible_text}\t{page_text}" if page_text else visible_text
    if not line_text:
        return None

    insert_at = doc.Content.End - 1
    paragraph_range = doc.Range(insert_at, insert_at)
    paragraph_range.InsertAfter(line_text)
    paragraph_range.InsertParagraphAfter()

    paragraph = doc.Range(insert_at, insert_at + len(line_text)).Paragraphs(1)
    normalize_toc_text_paragraph(
        paragraph,
        italic=italic,
        keep_with_next=keep_with_next,
        space_after=space_after,
        add_page_tab_stop=bool(page_text),
        doc=doc,
    )

    if italic and visible_text:
        text_range = doc.Range(insert_at, insert_at + len(visible_text))
        text_range.Font.Italic = True

    if page_text:
        page_start = insert_at + len(line_text) - len(page_text)
        page_range = doc.Range(page_start, page_start + len(page_text))
        page_range.Font.Name = "Times New Roman"
        page_range.Font.Size = 12
        page_range.Font.Bold = False
        page_range.Font.Italic = False

    return paragraph


def insert_special_toc_entry(doc, *, title: str, page: int | None) -> None:
    page_text = "" if page is None else str(page)
    insert_toc_text_paragraph(
        doc,
        text=title,
        italic=False,
        keep_with_next=False,
        space_after=TOC_SPECIAL_SPACE_AFTER,
        page_text=page_text,
    )


def insert_toc_entry(
    doc,
    *,
    title: str,
    authors: str,
    page: int | None,
    title_paragraph_index: int | None,
    is_special: bool = False,
) -> None:
    if is_special:
        insert_special_toc_entry(doc, title=title, page=page)
        return

    author_lines = wrap_author_lines(authors)
    page_text = "" if page is None else str(page)

    if not author_lines:
        if page_text:
            insert_special_toc_entry(doc, title=title, page=page)
            return

        insert_title_with_formatting(
            doc,
            title=title,
            title_paragraph_index=title_paragraph_index,
            keep_with_next=False,
            space_after=TOC_ARTICLE_SPACE_AFTER,
        )
        return

    keep_title_with_next = True
    insert_title_with_formatting(
        doc,
        title=title,
        title_paragraph_index=title_paragraph_index,
        keep_with_next=keep_title_with_next,
        space_after=TOC_INTERNAL_SPACE_AFTER,
    )

    for line_index, author_line in enumerate(author_lines, start=1):
        is_last_line = line_index == len(author_lines)
        insert_toc_text_paragraph(
            doc,
            text=author_line,
            italic=True,
            keep_with_next=not is_last_line,
            space_after=TOC_ARTICLE_SPACE_AFTER if is_last_line else TOC_INTERNAL_SPACE_AFTER,
            page_text=page_text if is_last_line else "",
        )


def main() -> int:
    args = parse_args()
    if win32 is None:
        raise RuntimeError("pywin32 is not installed. Install it with: pip install pywin32")

    run_tag = args.run_tag.strip()
    requested_source_path = Path(args.docx_path).resolve()
    source_path = resolve_final_insert_source_path(requested_source_path, run_tag).resolve()
    author_txt_path = resolve_output_path("author_index_ru_from_snapshot.txt", run_tag)
    output_docx_path = resolve_output_docx_path(run_tag)

    if not source_path.exists():
        raise FileNotFoundError(f"File not found: {source_path}")

    print(f"requested source docx path: {requested_source_path}")
    print(f"source docx path: {source_path}")
    print("loading TOC entries...")
    toc_entries = load_toc_entries(run_tag)
    print("loading author index lines...")
    author_lines = read_nonempty_lines(author_txt_path)

    output_docx_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"source docx: {source_path}")
    print(f"toc entries: {len(toc_entries)}")
    print(f"author lines: {len(author_lines)}")
    print(f"output docx path: {output_docx_path.resolve()}")

    print("opening Word...")
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    document = None
    try:
        document = word.Documents.Open(str(source_path))

        print("inserting TOC heading...")
        insert_page_break(document)
        insert_heading(document, "\u041e\u0433\u043b\u0430\u0432\u043b\u0435\u043d\u0438\u0435", level=1)
        insert_plain_paragraph(document, "")
        print("inserting TOC entries...")
        for index, entry in enumerate(toc_entries, start=1):
            insert_toc_entry(
                document,
                title=str(entry["title"]),
                authors=str(entry["authors"]),
                page=entry["page"] if isinstance(entry["page"], int) else None,
                title_paragraph_index=(
                    entry["title_paragraph_index"]
                    if isinstance(entry["title_paragraph_index"], int)
                    else None
                ),
                is_special=bool(entry.get("is_special")),
            )
            if index % TOC_PROGRESS_STEP == 0 or index == len(toc_entries):
                print(f"TOC entries inserted: {index}/{len(toc_entries)}")

        print("inserting author heading...")
        insert_page_break(document)
        insert_heading(document, "\u0410\u0432\u0442\u043e\u0440\u0441\u043a\u0438\u0439 \u0443\u043a\u0430\u0437\u0430\u0442\u0435\u043b\u044c", level=1)
        print("inserting author lines...")
        for index, line in enumerate(author_lines, start=1):
            insert_plain_paragraph(document, line)
            if index % AUTHOR_PROGRESS_STEP == 0 or index == len(author_lines):
                print(f"Author index lines inserted: {index}/{len(author_lines)}")

        print("saving output docx...")
        saved_output_path = save_document_with_fallback(document, output_docx_path)
        print(f"output docx: {saved_output_path}")
        print("done")
        return 0
    finally:
        if document is not None:
            document.Close(SaveChanges=False)
        word.Quit()


if __name__ == "__main__":
    raise SystemExit(main())

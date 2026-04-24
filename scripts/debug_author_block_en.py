from __future__ import annotations
import argparse
from dataclasses import dataclass
from pathlib import Path
import re
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import normalize_text, read_csv_rows, write_csv, write_text

try:
    import win32com.client as win32
except ImportError:
    win32 = None


INPUT_DIR = ROOT_DIR / "input"
OUTPUT_DIR = ROOT_DIR / "output"
DOCX_PATH = INPUT_DIR / "test1.docx"
INPUT_WINDOWS_CSV = OUTPUT_DIR / "author_windows_en_debug.csv"

OUTPUT_CSV = "author_block_en_debug.csv"
OUTPUT_TXT = "author_block_en_debug.txt"

KEYWORD_LABELS = ("keywords:", "key words:", "ketwords:")
TITLE_FONT_SIZE = 14.0
TITLE_FONT_TOLERANCE = 0.2
AUTHOR_LIST_WORD_LIMIT = 24
AFFILIATION_CUES = (
    "university",
    "institute",
    "department",
    "laboratory",
    "lab",
    "academy",
    "faculty",
    "centre",
    "center",
    "school",
    "college",
    "hospital",
    "museum",
    "company",
    "ltd",
    "inc",
    "llc",
    "corporation",
    "corp",
    "research center",
    "research centre",
    "research institute",
    "academy of",
    "institute of",
    "department of",
    "faculty of",
    "school of",
    "college of",
    "russia",
    "moscow",
    "kazan",
    "saint petersburg",
    "novosibirsk",
    "e-mail",
    "email",
)
SERVICE_PREFIXES = ("abstract", "annotation", "summary", "doi", "udc")
ADDRESS_LIKE_CUES = (
    "street",
    "st.",
    "avenue",
    "ave.",
    "road",
    "rd.",
    "boulevard",
    "blvd",
    "building",
    "room",
    "office",
    "postal",
    "zip code",
    "postbox",
    "p.o. box",
)
REFERENCE_STYLE_CUES = (
    "elsevier",
    "springer",
    "wiley",
    "publisher",
    "press",
    "amsterdam",
    "london",
    "new york",
    "vol.",
    "no.",
    " pp.",
    " p.",
    "//",
)
SERVICE_ROLE_CUES = (
    "student",
    "graduate student",
    "junior researcher",
    "postgraduate",
    "undergraduate",
    "phd student",
)
SHORT_INSTITUTION_FRAGMENTS = (
    "m.v. lomonosov",
    "a.v. topchiev",
    "d.i. mendeleev",
)

LATIN_LETTER_RE = re.compile(r"[A-Za-z]")
CYRILLIC_LETTER_RE = re.compile(r"[А-Яа-яЁё]")
LATIN_WORD_RE = re.compile(r"[A-Za-z][A-Za-z'-]*")

NAME_WORD = r"[A-Z][A-Za-z'-]+"
INITIAL_LETTER = r"(?:[A-Z]|[\u0410-\u042f\u0401\u0451])"
INITIAL_TAIL = r"(?:[a-z]|[\u0430-\u044f\u0451]){0,2}"
INITIAL = rf"{INITIAL_LETTER}{INITIAL_TAIL}\s*\."
DOUBLE_INITIALS = rf"{INITIAL}\s*{INITIAL}"
NAME_END = r"(?=[\W\d]|$)"
AUTHOR_PATTERNS = [
    re.compile(rf"\b{NAME_WORD}\s+{DOUBLE_INITIALS}{NAME_END}"),
    re.compile(rf"\b{DOUBLE_INITIALS}\s*{NAME_WORD}{NAME_END}"),
    re.compile(rf"\b{NAME_WORD}\s+{INITIAL}{NAME_END}"),
    re.compile(rf"\b{INITIAL}\s*{NAME_WORD}{NAME_END}"),
]
STRICT_AUTHOR_PATTERNS = AUTHOR_PATTERNS[:]


@dataclass
class ArticleWindowRow:
    article_no: int
    article_page: int | None
    ru_title_paragraph_text: str
    window_start_paragraph_index: int | None
    window_end_paragraph_index_exclusive: int | None
    status: str


@dataclass
class WindowLine:
    paragraph_index: int
    line_no: int
    text: str
    paragraph_first_char_font_size: float | None
    paragraph_first_char_italic: bool | None


@dataclass
class EnBlockDebugRow:
    article_no: int
    article_page: int | None
    ru_title_paragraph_text: str
    window_start_paragraph_index: int | None
    window_end_paragraph_index_exclusive: int | None
    keyword_paragraph_index: int | None
    keyword_line_text: str
    en_title_paragraph_index: int | None
    en_title: str
    author_paragraph_index: int | None
    raw_en_author_line: str
    font14_keyword_paragraph_index: int | None
    font14_keyword_line_text: str
    font14_title_paragraph_index: int | None
    font14_en_title: str
    font14_author_paragraph_index: int | None
    font14_raw_en_author_line: str
    font14_status: str
    source_used: str
    status: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build EN author block debug CSV/TXT from EN article windows."
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


def resolve_input_windows_csv(run_tag: str) -> Path:
    return resolve_output_path(INPUT_WINDOWS_CSV.name, run_tag) if run_tag.strip() else Path(INPUT_WINDOWS_CSV)


def parse_int(value: str) -> int | None:
    normalized = normalize_text(value)
    if normalized.isdigit():
        return int(normalized)
    return None


def split_nonempty_lines(text: str) -> list[str]:
    normalized = normalize_text(text)
    if not normalized:
        return []
    return [line for line in (normalize_text(item) for item in normalized.split("\n")) if line]


def read_article_windows(path: Path) -> list[ArticleWindowRow]:
    rows: list[ArticleWindowRow] = []

    for item in read_csv_rows(path):
        rows.append(
            ArticleWindowRow(
                article_no=parse_int(item.get("article_no", "")) or 0,
                article_page=parse_int(item.get("article_page", "")),
                ru_title_paragraph_text=normalize_text(item.get("ru_title_paragraph_text", "")),
                window_start_paragraph_index=parse_int(item.get("window_start_paragraph_index", "")),
                window_end_paragraph_index_exclusive=parse_int(item.get("window_end_paragraph_index_exclusive", "")),
                status=normalize_text(item.get("status", "")) or "missing_window",
            )
        )

    return rows


def build_window_lines(doc, start_index: int, end_exclusive: int) -> list[WindowLine]:
    lines: list[WindowLine] = []

    for paragraph_index in range(start_index, end_exclusive):
        paragraph_range = doc.Paragraphs(paragraph_index).Range
        text = normalize_text(paragraph_range.Text)
        if not text:
            continue

        paragraph_first_char_font_size = get_first_char_font_size(paragraph_range)
        paragraph_first_char_italic = get_first_char_italic(paragraph_range)

        for line_no, line in enumerate(split_nonempty_lines(text), start=1):
            lines.append(
                WindowLine(
                    paragraph_index=paragraph_index,
                    line_no=line_no,
                    text=line,
                    paragraph_first_char_font_size=paragraph_first_char_font_size,
                    paragraph_first_char_italic=paragraph_first_char_italic,
                )
            )

    return lines


def get_font_size_safe(range_obj) -> float | None:
    try:
        size = float(range_obj.Font.Size)
    except Exception:
        return None

    if size <= 0 or size > 100:
        return None
    return size


def get_first_char_font_size(paragraph_range) -> float | None:
    try:
        char_total = int(paragraph_range.Characters.Count)
    except Exception:
        return None

    for char_no in range(1, char_total + 1):
        try:
            char_range = paragraph_range.Characters(char_no)
            raw_char = char_range.Text
        except Exception:
            continue

        char = raw_char.replace("\r", "").replace("\n", "").replace("\x07", "").replace("\t", " ")
        if not char.strip():
            continue

        return get_font_size_safe(char_range)

    return None


def get_first_char_italic(paragraph_range) -> bool | None:
    try:
        char_total = int(paragraph_range.Characters.Count)
    except Exception:
        return None

    for char_no in range(1, char_total + 1):
        try:
            char_range = paragraph_range.Characters(char_no)
            raw_char = char_range.Text
        except Exception:
            continue

        char = raw_char.replace("\r", "").replace("\n", "").replace("\x07", "").replace("\t", " ")
        if not char.strip():
            continue

        try:
            italic_value = char_range.Font.Italic
        except Exception:
            return None

        try:
            return bool(int(italic_value))
        except Exception:
            return bool(italic_value)

    return None


def count_latin_letters(text: str) -> int:
    return sum(1 for char in text if char.isalpha() and char.isascii())


def count_cyrillic_letters(text: str) -> int:
    return len(CYRILLIC_LETTER_RE.findall(text))


def is_keyword_line(text: str) -> bool:
    normalized = normalize_text(text).casefold()
    return any(normalized.startswith(label) for label in KEYWORD_LABELS)


def is_service_line(text: str) -> bool:
    normalized = normalize_text(text).casefold()
    return any(normalized.startswith(prefix) for prefix in SERVICE_PREFIXES)


def has_affiliation_cues(text: str) -> bool:
    normalized = normalize_text(text)
    lowered = normalized.casefold()
    return any(cue in lowered for cue in AFFILIATION_CUES)


def looks_like_explicit_affiliation_line(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False

    words = LATIN_WORD_RE.findall(normalized)
    strict_matches = find_strict_author_matches(normalized)
    matches = find_author_matches(normalized)
    lowered = normalized.casefold()
    cue_hits = sum(1 for cue in AFFILIATION_CUES if cue in lowered)

    if "@" in normalized or "http://" in lowered or "https://" in lowered or "doi" in lowered or "//" in normalized:
        return True
    if any(cue in lowered for cue in ADDRESS_LIKE_CUES):
        return True
    if cue_hits == 0:
        return False
    if len(strict_matches) >= 1 and len(words) <= 10:
        return False
    if len(matches) >= 2 and len(words) <= AUTHOR_LIST_WORD_LIMIT:
        return False
    if len(matches) == 1 and len(words) <= 8 and ("," in normalized or "*" in normalized or any(char.isdigit() for char in normalized)):
        return False
    if cue_hits >= 2:
        return True
    if len(words) >= 10 and ("," in normalized or any(char.isdigit() for char in normalized)):
        return True
    if len(words) >= 12 and len(matches) == 0:
        return True
    return False


def looks_like_affiliation_line(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False

    words = LATIN_WORD_RE.findall(normalized)
    strict_matches = find_strict_author_matches(normalized)
    matches = find_author_matches(normalized)
    if looks_like_explicit_affiliation_line(normalized):
        return True
    if has_affiliation_cues(normalized):
        if len(strict_matches) >= 1 and len(words) <= 10:
            return False
        if len(matches) >= 2 and len(words) <= AUTHOR_LIST_WORD_LIMIT:
            return False
        if len(matches) == 1 and len(words) <= 7 and ("," in normalized or "*" in normalized or any(char.isdigit() for char in normalized)):
            return False
        return True

    if len(words) >= 7 and not strict_matches and len(matches) <= 1:
        return True

    if len(words) >= 9 and len(matches) <= 1 and "," in normalized:
        return True

    return False


def is_latin_rich(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False
    return count_latin_letters(normalized) >= 8 and count_cyrillic_letters(normalized) == 0 and bool(LATIN_LETTER_RE.search(normalized))


def is_latin_rich_author_text(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False
    if count_latin_letters(normalized) < 8 or not LATIN_LETTER_RE.search(normalized):
        return False
    return count_cyrillic_letters(normalized) == 0 or bool(find_author_matches(normalized))


def find_author_matches(text: str) -> list[str]:
    normalized = normalize_text(text)
    if not normalized:
        return []

    result: list[str] = []
    seen: set[str] = set()

    for pattern in AUTHOR_PATTERNS:
        for match in pattern.finditer(normalized):
            value = normalize_text(match.group(0)).strip(" ,;")
            key = value.casefold()
            if value and key not in seen:
                seen.add(key)
                result.append(value)

    return result


def find_strict_author_matches(text: str) -> list[str]:
    normalized = normalize_text(text)
    if not normalized:
        return []

    result: list[str] = []
    seen: set[str] = set()

    for pattern in STRICT_AUTHOR_PATTERNS:
        for match in pattern.finditer(normalized):
            value = normalize_text(match.group(0)).strip(" ,;")
            key = value.casefold()
            if value and key not in seen:
                seen.add(key)
                result.append(value)

    return result


def find_author_match_spans(text: str) -> list[tuple[int, int]]:
    normalized = normalize_text(text)
    if not normalized:
        return []

    candidates: list[tuple[int, int, int]] = []

    for priority, pattern in enumerate(AUTHOR_PATTERNS):
        for match in pattern.finditer(normalized):
            candidates.append((match.start(), match.end(), priority))

    candidates.sort(key=lambda item: (item[0], -(item[1] - item[0]), item[2]))

    selected: list[tuple[int, int]] = []
    for start, end, _ in candidates:
        if any(start < used_end and end > used_start for used_start, used_end in selected):
            continue
        selected.append((start, end))

    selected.sort()
    return selected


def looks_like_author_line(text: str) -> bool:
    normalized = normalize_text(text)
    if not is_latin_rich_author_text(normalized):
        return False
    if is_keyword_line(normalized) or is_service_line(normalized):
        return False
    if looks_like_bibliography_reference_line(normalized):
        return False
    if looks_like_affiliation_line(normalized):
        return False

    words = LATIN_WORD_RE.findall(normalized)
    if len(words) > 18:
        return False

    matches = find_author_matches(normalized)
    if len(matches) >= 2:
        return True
    if len(matches) == 1 and ("," in normalized or "*" in normalized or any(char.isdigit() for char in normalized) or len(words) <= 6):
        return True
    return False


def looks_like_title_line(text: str) -> bool:
    normalized = normalize_text(text)
    if not is_latin_rich(normalized):
        return False
    if is_keyword_line(normalized) or is_service_line(normalized) or looks_like_author_line(normalized):
        return False

    if looks_like_affiliation_line(normalized):
        return False
    if "@" in normalized:
        return False

    words = LATIN_WORD_RE.findall(normalized)
    if len(words) < 4 or len(words) > 30:
        return False
    if normalized.endswith("."):
        return False

    return True


def looks_like_bibliography_reference_line(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False

    words = LATIN_WORD_RE.findall(normalized)
    if len(words) < 5:
        return False

    lowered = normalized.casefold()
    has_leading_number = re.match(r"^\d+\.\s+", normalized) is not None
    has_reference_author_head = (
        re.match(rf"^\d+\.\s+{NAME_WORD}\s+{DOUBLE_INITIALS}", normalized) is not None
        or re.match(rf"^\d+\.\s+{NAME_WORD}\s+{INITIAL}", normalized) is not None
    )
    has_year = re.search(r"\b(?:19|20)\d{2}\b", normalized) is not None
    has_reference_cues = any(cue in lowered for cue in REFERENCE_STYLE_CUES)

    return has_leading_number and has_reference_author_head and (has_year or has_reference_cues)


def looks_like_backward_scan_noise(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False

    lowered = normalized.casefold()
    words = LATIN_WORD_RE.findall(normalized)
    matches = find_author_matches(normalized)
    strict_matches = find_strict_author_matches(normalized)
    compact = normalized.replace(" ", "")
    token_parts = re.findall(r"[A-Za-z0-9]+", normalized)

    if looks_like_bibliography_reference_line(normalized):
        return True
    if looks_like_affiliation_line(normalized):
        return True
    if "student" in lowered and ("year" in lowered or "specialist" in lowered or "program" in lowered):
        return True
    if re.fullmatch(r"(?:[A-Z][a-z]?(?:\s*,\s*|$)){4,}", normalized):
        return True

    if len(matches) == 0 and len(strict_matches) == 0:
        if re.fullmatch(r"[A-Za-z0-9()/+\-.,*]+", compact) and sum(1 for char in compact if char.isdigit()) >= 2:
            return True
        if any(char.isdigit() for char in normalized) and any(char in normalized for char in "/()=+-") and len(words) <= 8:
            return True
        if 1 <= len(token_parts) <= 6:
            noisy_tokens = sum(
                1 for part in token_parts if any(char.isdigit() for char in part) or part.isupper()
            )
            if noisy_tokens >= max(1, len(token_parts) - 1):
                return True

    return False


def has_title_font_size(line: WindowLine) -> bool:
    if line.paragraph_first_char_font_size is None:
        return False
    return abs(line.paragraph_first_char_font_size - TITLE_FONT_SIZE) <= TITLE_FONT_TOLERANCE


def looks_like_post_title_author_line(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False
    if not is_latin_rich_author_text(normalized):
        return False
    if is_keyword_line(normalized) or is_service_line(normalized):
        return False

    if looks_like_bibliography_reference_line(normalized):
        return False
    if looks_like_explicit_affiliation_line(normalized):
        return False
    if re.match(r"^\d+\.", normalized):
        return False

    words = LATIN_WORD_RE.findall(normalized)
    strict_matches = find_strict_author_matches(normalized)
    matches = find_author_matches(normalized)
    if len(matches) >= 2:
        return True
    if len(strict_matches) >= 1 and len(words) <= 10:
        return True
    if len(matches) == 1 and (
        "," in normalized
        or "*" in normalized
        or any(char.isdigit() for char in normalized)
        or len(words) <= 8
    ):
        return True
    return False


def has_structural_author_font_signal(line: WindowLine, title_line: WindowLine) -> bool:
    if line.paragraph_index == title_line.paragraph_index:
        return True
    if line.paragraph_first_char_font_size is None or title_line.paragraph_first_char_font_size is None:
        return False
    return line.paragraph_first_char_font_size < title_line.paragraph_first_char_font_size - TITLE_FONT_TOLERANCE


def has_non_italic_author_style_signal(line: WindowLine, title_line: WindowLine) -> bool:
    if line.paragraph_index == title_line.paragraph_index:
        return True
    if line.paragraph_first_char_italic is None:
        return True
    return not line.paragraph_first_char_italic


def has_author_like_fragments(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False

    words = LATIN_WORD_RE.findall(normalized)
    matches = find_author_matches(normalized)
    strict_matches = find_strict_author_matches(normalized)

    if len(matches) >= 1 or len(strict_matches) >= 1:
        return True
    if re.search(rf"\b{INITIAL}", normalized):
        return True
    if len(words) <= 8 and ("," in normalized or "*" in normalized or any(char.isdigit() for char in normalized)):
        return True
    return False


def looks_like_structural_post_title_author_start(line: WindowLine, title_line: WindowLine) -> bool:
    normalized = normalize_text(line.text)
    if not normalized:
        return False
    if is_keyword_line(normalized) or is_service_line(normalized):
        return False
    if not is_latin_rich_author_text(normalized):
        return False
    if re.match(r"^\d+\.", normalized):
        return False
    if not has_author_like_fragments(normalized):
        return False

    words = LATIN_WORD_RE.findall(normalized)
    if line.paragraph_index == title_line.paragraph_index:
        if looks_like_explicit_affiliation_line(normalized):
            return False
        if looks_like_post_title_author_line(normalized):
            return True
        return len(words) <= 14
    if line.paragraph_index > title_line.paragraph_index:
        if not has_non_italic_author_style_signal(line, title_line):
            return False
        if looks_like_explicit_affiliation_line(normalized):
            return False
        if looks_like_post_title_author_line(normalized):
            return True
        if has_structural_author_font_signal(line, title_line):
            return len(words) <= 14
        return len(words) <= 10
    return False


def looks_like_author_list_continuation(text: str, previous_text: str = "") -> bool:
    normalized = normalize_text(text)
    previous_normalized = normalize_text(previous_text)
    if not normalized:
        return False
    if not is_latin_rich_author_text(normalized):
        return False
    if is_keyword_line(normalized) or is_service_line(normalized):
        return False
    if looks_like_affiliation_line(normalized):
        return False

    words = LATIN_WORD_RE.findall(normalized)
    strict_matches = find_strict_author_matches(normalized)
    matches = find_author_matches(normalized)

    if len(matches) >= 2:
        return True
    if len(strict_matches) >= 1 and len(words) <= 8:
        return True
    if len(matches) == 1 and len(words) <= 8:
        return True
    if previous_normalized.endswith((",", ";")) and len(words) <= 2:
        if re.fullmatch(rf"{NAME_WORD}(?:\s+{NAME_WORD})?", normalized):
            return True
    if has_author_like_fragments(normalized) and len(words) <= 8 and not looks_like_affiliation_line(normalized):
        return True
    return False


def looks_like_service_or_affiliation_only_author_candidate(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return True

    lowered = normalized.casefold()
    compact_lowered = lowered.strip(" ,;*")
    matches = find_author_matches(normalized)
    strict_matches = find_strict_author_matches(normalized)

    if is_keyword_line(normalized) or is_service_line(normalized):
        return True
    if looks_like_bibliography_reference_line(normalized):
        return True
    if compact_lowered in SHORT_INSTITUTION_FRAGMENTS:
        return True
    if "@" in normalized or "http://" in lowered or "https://" in lowered or "//" in normalized:
        return True
    if any(cue in lowered for cue in ADDRESS_LIKE_CUES):
        return True
    if looks_like_explicit_affiliation_line(normalized) or looks_like_affiliation_line(normalized):
        return True
    if any(cue in lowered for cue in SERVICE_ROLE_CUES) and len(matches) <= 1 and len(strict_matches) <= 1:
        return True
    return False


def has_real_author_candidate(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False
    if looks_like_service_or_affiliation_only_author_candidate(normalized):
        return False
    return looks_like_author_line(normalized) or looks_like_post_title_author_line(normalized)


def finalize_title_author_candidates(
    title_line: WindowLine | None,
    author_line: WindowLine | None,
) -> tuple[WindowLine | None, WindowLine | None]:
    split_author_line: WindowLine | None = None

    if title_line is not None:
        split_title_line, split_author_line = split_combined_title_author_line(title_line)
        if split_author_line is not None:
            title_line = split_title_line

    if author_line is not None:
        author_line = trim_author_line_tail(author_line)
        if not has_real_author_candidate(author_line.text):
            author_line = None

    if author_line is None and split_author_line is not None and has_real_author_candidate(split_author_line.text):
        author_line = split_author_line

    return title_line, author_line


def choose_from_keyword_window(lines: list[WindowLine]) -> tuple[WindowLine | None, WindowLine | None, WindowLine | None]:
    keyword_line: WindowLine | None = None

    for line in lines:
        if is_keyword_line(line.text):
            keyword_line = line
            break

    keyword_index = lines.index(keyword_line) if keyword_line is not None else len(lines)
    structural_title_line, structural_author_line = choose_structural_en_block(lines, keyword_index)
    if structural_title_line is not None and structural_author_line is not None:
        return keyword_line, structural_title_line, structural_author_line

    author_line: WindowLine | None = None
    for line in reversed(lines[:keyword_index]):
        if looks_like_author_line(line.text) and not looks_like_backward_scan_noise(line.text):
            author_line = line
            break

    title_anchor = lines.index(author_line) if author_line is not None else keyword_index

    title_line: WindowLine | None = None
    for line in reversed(lines[:title_anchor]):
        if looks_like_title_line(line.text) and not looks_like_backward_scan_noise(line.text):
            title_line = line
            break

    if title_line is None or author_line is None:
        fallback_title_line, fallback_author_line = choose_by_keyword_proximity(lines, keyword_index)
        if fallback_author_line is not None:
            author_line = fallback_author_line
            if fallback_title_line is not None:
                title_line = fallback_title_line

    if title_line is None and structural_title_line is not None:
        title_line = structural_title_line
    if author_line is None and structural_author_line is not None:
        author_line = structural_author_line

    title_line, author_line = finalize_title_author_candidates(title_line, author_line)

    return keyword_line, title_line, author_line


def choose_by_forward_scan(lines: list[WindowLine]) -> tuple[WindowLine | None, WindowLine | None]:
    structural_title_line, structural_author_line = choose_structural_en_block(lines, len(lines))
    if structural_title_line is not None and structural_author_line is not None:
        return structural_title_line, structural_author_line

    title_line: WindowLine | None = structural_title_line

    for line in lines:
        if looks_like_title_line(line.text):
            title_line = line
            break

    if title_line is None:
        return None, None

    title_index = lines.index(title_line)
    author_line: WindowLine | None = None

    for line in lines[title_index + 1 :]:
        if looks_like_author_line(line.text):
            author_line = line
            break
        if is_keyword_line(line.text):
            break

    if author_line is None and structural_author_line is not None:
        author_line = structural_author_line

    return finalize_title_author_candidates(title_line, author_line)


def make_window_line(base_line: WindowLine, text: str) -> WindowLine:
    return WindowLine(
        paragraph_index=base_line.paragraph_index,
        line_no=base_line.line_no,
        text=normalize_text(text),
        paragraph_first_char_font_size=base_line.paragraph_first_char_font_size,
        paragraph_first_char_italic=base_line.paragraph_first_char_italic,
    )


def join_window_line_texts(lines: list[WindowLine]) -> str:
    return normalize_text(" ".join(normalize_text(line.text) for line in lines if normalize_text(line.text)))


def find_window_line_index(lines: list[WindowLine], target_line: WindowLine) -> int | None:
    for index, line in enumerate(lines):
        if line.paragraph_index == target_line.paragraph_index and line.line_no == target_line.line_no:
            return index
    return None


def split_combined_title_author_line(line: WindowLine) -> tuple[WindowLine, WindowLine | None]:
    normalized = normalize_text(line.text)
    if not normalized:
        return line, None

    candidates: list[tuple[int, str, str]] = []

    for pattern in AUTHOR_PATTERNS:
        for match in pattern.finditer(normalized):
            start = match.start()
            if start < int(len(normalized) * 0.35):
                continue

            title_part = normalize_text(normalized[:start]).strip(" ,;")
            author_part = normalize_text(normalized[start:]).strip(" ,;")
            title_words = LATIN_WORD_RE.findall(title_part)
            author_words = LATIN_WORD_RE.findall(author_part)

            if len(title_words) < 4 or len(author_words) == 0 or len(author_words) > AUTHOR_LIST_WORD_LIMIT:
                continue
            if not looks_like_split_title_part(title_part, line):
                continue
            if not looks_like_post_title_author_line(author_part):
                continue

            candidates.append((start, title_part, author_part))

    if not candidates:
        return line, None

    _, title_part, author_part = min(candidates, key=lambda item: item[0])
    return make_window_line(line, title_part), trim_author_line_tail(make_window_line(line, author_part))


def looks_like_split_title_part(text: str, source_line: WindowLine) -> bool:
    normalized = normalize_text(text)
    if looks_like_title_line(normalized):
        return True
    if not has_title_font_size(source_line):
        return False
    if not is_latin_rich(normalized):
        return False
    if is_keyword_line(normalized) or is_service_line(normalized):
        return False
    if "@" in normalized:
        return False

    words = LATIN_WORD_RE.findall(normalized)
    if len(words) < 4 or len(words) > 30:
        return False
    if normalized.endswith("."):
        return False
    return True


def trim_author_line_tail(line: WindowLine) -> WindowLine:
    normalized = normalize_text(line.text)
    if not normalized:
        return line

    spans = find_author_match_spans(normalized)
    if not spans:
        return line

    last_end = spans[0][1]
    for next_start, next_end in spans[1:]:
        gap_text = normalize_text(normalized[last_end:next_start])
        lowered_gap = gap_text.casefold()
        if gap_text and (
            "student" in lowered_gap
            or any(cue in lowered_gap for cue in AFFILIATION_CUES)
            or any(cue in lowered_gap for cue in ADDRESS_LIKE_CUES)
            or "@" in gap_text
            or "http://" in lowered_gap
            or "https://" in lowered_gap
            or "//" in gap_text
            or is_service_line(gap_text)
        ):
            break
        last_end = next_end

    trim_probe = last_end
    while trim_probe < len(normalized) and normalized[trim_probe] in " ,;*0123456789":
        trim_probe += 1

    tail = normalize_text(normalized[trim_probe:])
    if not tail:
        return make_window_line(line, normalized[:trim_probe].rstrip(" ,;"))

    lowered_tail = tail.casefold()
    if (
        "student" in lowered_tail
        or any(cue in lowered_tail for cue in AFFILIATION_CUES)
        or any(cue in lowered_tail for cue in ADDRESS_LIKE_CUES)
        or "@" in tail
        or "http://" in lowered_tail
        or "https://" in lowered_tail
        or "//" in tail
        or is_service_line(tail)
    ):
        return make_window_line(line, normalized[:trim_probe].rstrip(" ,;"))

    return line


def choose_by_keyword_proximity(lines: list[WindowLine], keyword_index: int) -> tuple[WindowLine | None, WindowLine | None]:
    paragraph_blocks: list[WindowLine] = []
    current_index = keyword_index - 1

    while current_index >= 0:
        paragraph_index = lines[current_index].paragraph_index
        paragraph_end = current_index + 1

        while current_index >= 0 and lines[current_index].paragraph_index == paragraph_index:
            current_index -= 1

        paragraph_start = current_index + 1
        paragraph_blocks.append(
            make_window_line(lines[paragraph_start], join_window_line_texts(lines[paragraph_start:paragraph_end]))
        )

    author_line: WindowLine | None = None
    title_line: WindowLine | None = None
    author_block_no: int | None = None

    for block_no, paragraph_line in enumerate(paragraph_blocks):
        normalized = normalize_text(paragraph_line.text)
        if not normalized:
            continue
        if is_keyword_line(normalized) or is_service_line(normalized):
            continue
        if looks_like_explicit_affiliation_line(normalized) or looks_like_backward_scan_noise(normalized):
            continue

        split_title_line, split_author_line = split_combined_title_author_line(paragraph_line)
        if split_author_line is not None:
            return split_title_line, split_author_line

        words = LATIN_WORD_RE.findall(normalized)
        if looks_like_author_line(normalized) or looks_like_post_title_author_line(normalized):
            author_line = trim_author_line_tail(paragraph_line)
            author_block_no = block_no
            break
        if has_author_like_fragments(normalized) and len(words) <= 14:
            author_line = trim_author_line_tail(paragraph_line)
            author_block_no = block_no
            break

    if author_line is None or author_block_no is None:
        return None, None

    for paragraph_line in paragraph_blocks[author_block_no + 1 :]:
        normalized = normalize_text(paragraph_line.text)
        if not normalized:
            continue
        if is_keyword_line(normalized) or is_service_line(normalized):
            continue
        if looks_like_explicit_affiliation_line(normalized) or looks_like_backward_scan_noise(normalized):
            continue

        words = LATIN_WORD_RE.findall(normalized)
        if looks_like_title_line(normalized):
            title_line = paragraph_line
            break
        if is_latin_rich(normalized) and not has_author_like_fragments(normalized) and 4 <= len(words) <= 30:
            if not normalized.endswith("."):
                title_line = paragraph_line
                break

    return title_line, author_line


def has_recent_cyrillic_context(lines: list[WindowLine], line_index: int) -> bool:
    start_index = max(0, line_index - 40)

    for candidate in reversed(lines[start_index:line_index]):
        normalized = normalize_text(candidate.text)
        if not normalized:
            continue
        if count_cyrillic_letters(normalized) >= 6:
            return True

    return False


def has_strong_title_before_author_pattern(lines: list[WindowLine], line_index: int) -> bool:
    line = lines[line_index]
    normalized = normalize_text(line.text)
    if not normalized:
        return False
    if line.paragraph_first_char_italic is True:
        return False
    if not is_latin_rich(normalized):
        return False
    if not has_title_font_size(line):
        return False
    if is_keyword_line(normalized) or is_service_line(normalized):
        return False
    if looks_like_author_line(normalized) or looks_like_post_title_author_line(normalized):
        return False
    if line_index + 1 >= len(lines):
        return False
    return looks_like_structural_post_title_author_start(lines[line_index + 1], line)


def should_reject_structural_title_as_affiliation(
    lines: list[WindowLine],
    line_index: int,
) -> bool:
    line = lines[line_index]
    normalized = normalize_text(line.text)
    if not normalized:
        return False
    if has_strong_title_before_author_pattern(lines, line_index):
        return False
    if looks_like_explicit_affiliation_line(normalized):
        return True
    if not looks_like_affiliation_line(normalized):
        return False
    if has_affiliation_cues(normalized):
        return True
    if has_author_like_fragments(normalized):
        return True
    return False


def looks_like_structural_en_title_candidate(lines: list[WindowLine], line_index: int) -> bool:
    line = lines[line_index]
    normalized = normalize_text(line.text)
    if not normalized:
        return False
    if line.paragraph_first_char_italic is True:
        return False
    if not is_latin_rich(normalized):
        return False
    if is_keyword_line(normalized) or is_service_line(normalized):
        return False
    if has_strong_title_before_author_pattern(lines, line_index):
        return True
    if should_reject_structural_title_as_affiliation(lines, line_index):
        return False
    if looks_like_author_line(normalized) or looks_like_post_title_author_line(normalized):
        return False

    words = LATIN_WORD_RE.findall(normalized)
    if len(words) < 4 or len(words) > 24:
        return False
    if has_author_like_fragments(normalized) and len(words) <= 12:
        return False

    return has_title_font_size(line) or looks_like_title_line(normalized)


def expand_title_line(lines: list[WindowLine], title_index: int) -> WindowLine:
    base_line = lines[title_index]
    start_index = title_index
    end_index = title_index

    while start_index > 0:
        candidate = lines[start_index - 1]
        if candidate.paragraph_index != base_line.paragraph_index:
            break
        if looks_like_author_line(candidate.text) or looks_like_post_title_author_line(candidate.text):
            break
        if not is_latin_rich(candidate.text) or is_keyword_line(candidate.text) or is_service_line(candidate.text):
            break
        start_index -= 1

    while end_index + 1 < len(lines):
        candidate = lines[end_index + 1]
        if candidate.paragraph_index != base_line.paragraph_index:
            break
        if candidate.paragraph_first_char_italic is True:
            break
        if is_keyword_line(candidate.text) or is_service_line(candidate.text):
            break
        if looks_like_author_line(candidate.text) or looks_like_post_title_author_line(candidate.text):
            break
        if should_reject_structural_title_as_affiliation(lines, end_index + 1):
            break
        if not is_latin_rich(candidate.text):
            break
        if not has_title_font_size(candidate) and not looks_like_title_line(candidate.text):
            break
        end_index += 1

    return make_window_line(base_line, join_window_line_texts(lines[start_index : end_index + 1]))


def choose_structural_en_block(
    lines: list[WindowLine],
    keyword_index: int,
) -> tuple[WindowLine | None, WindowLine | None]:
    title_only_line: WindowLine | None = None

    for title_index in range(keyword_index):
        candidate = lines[title_index]
        if not looks_like_structural_en_title_candidate(lines, title_index):
            continue
        if not has_recent_cyrillic_context(lines, title_index):
            continue

        expanded_title_line = expand_title_line(lines, title_index)
        author_line = collect_same_paragraph_author_line(lines, title_index)

        if author_line is None:
            author_line = collect_following_paragraph_author_line(lines, title_index, keyword_index)

        if author_line is None:
            expanded_title_line, author_line = split_combined_title_author_line(expanded_title_line)

        expanded_title_line, author_line = finalize_title_author_candidates(expanded_title_line, author_line)

        if title_only_line is None and expanded_title_line is not None:
            title_only_line = expanded_title_line
        if expanded_title_line is not None and author_line is not None:
            return expanded_title_line, author_line

    return title_only_line, None


def collect_same_paragraph_author_line(lines: list[WindowLine], title_index: int) -> WindowLine | None:
    title_line = lines[title_index]
    author_start_index: int | None = None

    for index in range(title_index + 1, len(lines)):
        candidate = lines[index]
        if candidate.paragraph_index != title_line.paragraph_index:
            break
        if looks_like_structural_post_title_author_start(candidate, title_line):
            author_start_index = index
            break

    if author_start_index is None:
        return None

    author_end_index = author_start_index
    for index in range(author_start_index + 1, len(lines)):
        candidate = lines[index]
        if candidate.paragraph_index != title_line.paragraph_index:
            break
        if is_service_line(candidate.text) or is_keyword_line(candidate.text):
            break
        if not looks_like_author_list_continuation(candidate.text, lines[author_end_index].text):
            break
        author_end_index = index

    return trim_author_line_tail(
        make_window_line(
            lines[author_start_index],
            join_window_line_texts(lines[author_start_index : author_end_index + 1]),
        )
    )


def collect_following_paragraph_author_line(
    lines: list[WindowLine],
    title_index: int,
    keyword_index: int,
) -> WindowLine | None:
    title_line = lines[title_index]
    current_index = title_index + 1

    while current_index < keyword_index and lines[current_index].paragraph_index == title_line.paragraph_index:
        current_index += 1

    if current_index >= keyword_index:
        return None

    paragraph_index = lines[current_index].paragraph_index
    paragraph_lines: list[WindowLine] = []
    while current_index < keyword_index and lines[current_index].paragraph_index == paragraph_index:
        paragraph_lines.append(lines[current_index])
        current_index += 1

    paragraph_text = join_window_line_texts(paragraph_lines)
    if not paragraph_text or is_keyword_line(paragraph_text):
        return None
    if is_service_line(paragraph_text):
        return None

    author_start_index: int | None = None
    for index, line in enumerate(paragraph_lines):
        if looks_like_structural_post_title_author_start(line, title_line):
            author_start_index = index
            break

    if author_start_index is None:
        return None

    author_end_index = author_start_index
    for index in range(author_start_index + 1, len(paragraph_lines)):
        if not looks_like_author_list_continuation(paragraph_lines[index].text, paragraph_lines[author_end_index].text):
            break
        author_end_index = index

    return trim_author_line_tail(
        make_window_line(
            paragraph_lines[author_start_index],
            join_window_line_texts(paragraph_lines[author_start_index : author_end_index + 1]),
        )
    )


def choose_by_keyword_and_font14(
    lines: list[WindowLine],
) -> tuple[WindowLine | None, WindowLine | None, WindowLine | None]:
    keyword_line: WindowLine | None = None

    for line in lines:
        if is_keyword_line(line.text):
            keyword_line = line
            break

    keyword_index = lines.index(keyword_line) if keyword_line is not None else len(lines)

    structural_title_line, structural_author_line = choose_structural_en_block(lines, keyword_index)
    if structural_title_line is not None and structural_author_line is not None:
        return keyword_line, structural_title_line, structural_author_line

    title_line = structural_title_line

    for line in reversed(lines[:keyword_index]):
        if has_title_font_size(line) and looks_like_title_line(line.text):
            title_line = line
            break

    if title_line is None:
        return keyword_line, structural_title_line, structural_author_line

    title_index = find_window_line_index(lines, title_line)
    if title_index is None:
        return keyword_line, structural_title_line, structural_author_line

    expanded_title_line = expand_title_line(lines, title_index)
    author_line = collect_same_paragraph_author_line(lines, title_index)

    if author_line is None:
        author_line = collect_following_paragraph_author_line(lines, title_index, keyword_index)

    if author_line is None:
        expanded_title_line, author_line = split_combined_title_author_line(expanded_title_line)

    expanded_title_line, author_line = finalize_title_author_candidates(expanded_title_line, author_line)

    if expanded_title_line is None and structural_title_line is not None:
        expanded_title_line = structural_title_line
    if author_line is None and structural_author_line is not None:
        author_line = structural_author_line

    return keyword_line, expanded_title_line, author_line


def build_debug_row(doc, window: ArticleWindowRow, paragraph_total: int) -> EnBlockDebugRow:
    if window.window_start_paragraph_index is None:
        return EnBlockDebugRow(
            article_no=window.article_no,
            article_page=window.article_page,
            ru_title_paragraph_text=window.ru_title_paragraph_text,
            window_start_paragraph_index=window.window_start_paragraph_index,
            window_end_paragraph_index_exclusive=window.window_end_paragraph_index_exclusive,
            keyword_paragraph_index=None,
            keyword_line_text="",
            en_title_paragraph_index=None,
            en_title="",
            author_paragraph_index=None,
            raw_en_author_line="",
            font14_keyword_paragraph_index=None,
            font14_keyword_line_text="",
            font14_title_paragraph_index=None,
            font14_en_title="",
            font14_author_paragraph_index=None,
            font14_raw_en_author_line="",
            font14_status="missing_window",
            source_used="not_found",
            status="missing_window",
        )

    end_exclusive = window.window_end_paragraph_index_exclusive or (paragraph_total + 1)
    lines = build_window_lines(doc, window.window_start_paragraph_index, min(end_exclusive, paragraph_total + 1))

    keyword_line, title_line, author_line = choose_from_keyword_window(lines)
    source_used = "keyword_backward_scan" if keyword_line is not None else "forward_scan"

    font14_keyword_line, font14_title_line, font14_author_line = choose_by_keyword_and_font14(lines)

    if title_line is None and author_line is None:
        fallback_title, fallback_author = choose_by_forward_scan(lines)
        title_line = fallback_title
        author_line = fallback_author

    if title_line is not None and author_line is not None:
        status = "found_title_and_author"
    elif title_line is not None:
        status = "found_title_only"
    else:
        status = "not_found"
        source_used = "not_found" if keyword_line is None else source_used

    if font14_title_line is not None and font14_author_line is not None:
        font14_status = "found_title_and_author"
    elif font14_title_line is not None:
        font14_status = "found_title_only"
    elif font14_keyword_line is not None:
        font14_status = "keyword_found_no_font14_title"
    else:
        font14_status = "keyword_not_found"

    return EnBlockDebugRow(
        article_no=window.article_no,
        article_page=window.article_page,
        ru_title_paragraph_text=window.ru_title_paragraph_text,
        window_start_paragraph_index=window.window_start_paragraph_index,
        window_end_paragraph_index_exclusive=window.window_end_paragraph_index_exclusive,
        keyword_paragraph_index=keyword_line.paragraph_index if keyword_line is not None else None,
        keyword_line_text=keyword_line.text if keyword_line is not None else "",
        en_title_paragraph_index=title_line.paragraph_index if title_line is not None else None,
        en_title=title_line.text if title_line is not None else "",
        author_paragraph_index=author_line.paragraph_index if author_line is not None else None,
        raw_en_author_line=author_line.text if author_line is not None else "",
        font14_keyword_paragraph_index=font14_keyword_line.paragraph_index if font14_keyword_line is not None else None,
        font14_keyword_line_text=font14_keyword_line.text if font14_keyword_line is not None else "",
        font14_title_paragraph_index=font14_title_line.paragraph_index if font14_title_line is not None else None,
        font14_en_title=font14_title_line.text if font14_title_line is not None else "",
        font14_author_paragraph_index=font14_author_line.paragraph_index if font14_author_line is not None else None,
        font14_raw_en_author_line=font14_author_line.text if font14_author_line is not None else "",
        font14_status=font14_status,
        source_used=source_used,
        status=status,
    )


def build_debug_rows(docx_path: Path, windows: list[ArticleWindowRow]) -> list[EnBlockDebugRow]:
    if win32 is None:
        raise RuntimeError("Не установлен pywin32. Установите: pip install pywin32")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    doc = None

    try:
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        paragraph_total = int(doc.Paragraphs.Count)
        rows: list[EnBlockDebugRow] = []

        for index, window in enumerate(windows, start=1):
            if index == 1 or index % 5 == 0 or index == len(windows):
                window_start = window.window_start_paragraph_index if window.window_start_paragraph_index is not None else "?"
                window_end = window.window_end_paragraph_index_exclusive if window.window_end_paragraph_index_exclusive is not None else "<document_end>"
                page_text = window.article_page if window.article_page is not None else "?"
                print(
                    f"[progress] article {index}/{len(windows)} "
                    f"(article_no={window.article_no}, page={page_text}, window={window_start}->{window_end})"
                )

            rows.append(build_debug_row(doc, window, paragraph_total))

        return rows
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


def write_debug_csv(path: Path, rows: list[EnBlockDebugRow]) -> None:
    csv_rows = (
        [
            row.article_no,
            row.article_page if row.article_page is not None else "",
            row.ru_title_paragraph_text,
            row.window_start_paragraph_index if row.window_start_paragraph_index is not None else "",
            row.window_end_paragraph_index_exclusive if row.window_end_paragraph_index_exclusive is not None else "",
            row.keyword_paragraph_index if row.keyword_paragraph_index is not None else "",
            row.keyword_line_text,
            row.en_title_paragraph_index if row.en_title_paragraph_index is not None else "",
            row.en_title,
            row.author_paragraph_index if row.author_paragraph_index is not None else "",
            row.raw_en_author_line,
            row.font14_keyword_paragraph_index if row.font14_keyword_paragraph_index is not None else "",
            row.font14_keyword_line_text,
            row.font14_title_paragraph_index if row.font14_title_paragraph_index is not None else "",
            row.font14_en_title,
            row.font14_author_paragraph_index if row.font14_author_paragraph_index is not None else "",
            row.font14_raw_en_author_line,
            row.font14_status,
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
            "ru_title_paragraph_text",
            "window_start_paragraph_index",
            "window_end_paragraph_index_exclusive",
            "keyword_paragraph_index",
            "keyword_line_text",
            "en_title_paragraph_index",
            "en_title",
            "author_paragraph_index",
            "raw_en_author_line",
            "font14_keyword_paragraph_index",
            "font14_keyword_line_text",
            "font14_title_paragraph_index",
            "font14_en_title",
            "font14_author_paragraph_index",
            "font14_raw_en_author_line",
            "font14_status",
            "source_used",
            "status",
        ],
        csv_rows,
    )


def build_debug_text(rows: list[EnBlockDebugRow]) -> str:
    blocks: list[str] = []

    for row in rows:
        blocks.append(
            "\n".join(
                [
                    "=" * 100,
                    f"ARTICLE {row.article_no} | page {row.article_page if row.article_page is not None else '?'}",
                    f"WINDOW: {row.window_start_paragraph_index if row.window_start_paragraph_index is not None else '?'} -> {row.window_end_paragraph_index_exclusive if row.window_end_paragraph_index_exclusive is not None else '<document_end>'}",
                    "RU TITLE PARAGRAPH:",
                    row.ru_title_paragraph_text or "<empty>",
                    "",
                    f"KEYWORD PARAGRAPH INDEX: {row.keyword_paragraph_index if row.keyword_paragraph_index is not None else '?'}",
                    "KEYWORD LINE:",
                    row.keyword_line_text or "<not found>",
                    "",
                    f"EN TITLE PARAGRAPH INDEX: {row.en_title_paragraph_index if row.en_title_paragraph_index is not None else '?'}",
                    "EN TITLE:",
                    row.en_title or "<not found>",
                    "",
                    f"AUTHOR PARAGRAPH INDEX: {row.author_paragraph_index if row.author_paragraph_index is not None else '?'}",
                    "RAW EN AUTHOR LINE:",
                    row.raw_en_author_line or "<not found>",
                    "",
                    f"FONT14 KEYWORD PARAGRAPH INDEX: {row.font14_keyword_paragraph_index if row.font14_keyword_paragraph_index is not None else '?'}",
                    "FONT14 KEYWORD LINE:",
                    row.font14_keyword_line_text or "<not found>",
                    "",
                    f"FONT14 TITLE PARAGRAPH INDEX: {row.font14_title_paragraph_index if row.font14_title_paragraph_index is not None else '?'}",
                    "FONT14 EN TITLE:",
                    row.font14_en_title or "<not found>",
                    "",
                    f"FONT14 AUTHOR PARAGRAPH INDEX: {row.font14_author_paragraph_index if row.font14_author_paragraph_index is not None else '?'}",
                    "FONT14 RAW EN AUTHOR LINE:",
                    row.font14_raw_en_author_line or "<not found>",
                    "",
                    f"FONT14 STATUS: {row.font14_status}",
                    "",
                    f"SOURCE USED: {row.source_used}",
                    f"STATUS: {row.status}",
                ]
            )
        )

    return "\n\n".join(blocks) + ("\n" if blocks else "")


def print_summary(rows: list[EnBlockDebugRow]) -> None:
    title_and_author = sum(1 for row in rows if row.status == "found_title_and_author")
    title_only = sum(1 for row in rows if row.status == "found_title_only")
    not_found = sum(1 for row in rows if row.status == "not_found")
    missing_window = sum(1 for row in rows if row.status == "missing_window")
    font14_title_and_author = sum(1 for row in rows if row.font14_status == "found_title_and_author")
    font14_title_only = sum(1 for row in rows if row.font14_status == "found_title_only")
    font14_no_title = sum(1 for row in rows if row.font14_status == "keyword_found_no_font14_title")
    font14_no_keyword = sum(1 for row in rows if row.font14_status == "keyword_not_found")

    print("=" * 100)
    print(f"ARTICLE WINDOWS READ:    {len(rows)}")
    print(f"TITLE AND AUTHOR FOUND:  {title_and_author}")
    print(f"TITLE ONLY FOUND:        {title_only}")
    print(f"NOT FOUND:               {not_found}")
    print(f"MISSING WINDOW:          {missing_window}")
    print("-" * 100)
    print(f"FONT14 TITLE+AUTHOR:     {font14_title_and_author}")
    print(f"FONT14 TITLE ONLY:       {font14_title_only}")
    print(f"FONT14 NO TITLE:         {font14_no_title}")
    print(f"FONT14 NO KEYWORD:       {font14_no_keyword}")
    print("=" * 100)


def main() -> None:
    args = parse_args()
    run_tag = args.run_tag.strip()
    input_path = resolve_input_windows_csv(run_tag)
    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"Файл не найден: {docx_path}")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    windows = read_article_windows(input_path)
    rows = build_debug_rows(docx_path, windows)

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

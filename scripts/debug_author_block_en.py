from __future__ import annotations

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
AFFILIATION_CUES = (
    "university",
    "institute",
    "department",
    "laboratory",
    "academy",
    "faculty",
    "college",
    "russia",
    "moscow",
    "kazan",
    "saint petersburg",
    "novosibirsk",
    "e-mail",
    "email",
)
SERVICE_PREFIXES = ("abstract", "annotation", "summary", "doi", "udc")

LATIN_LETTER_RE = re.compile(r"[A-Za-z]")
CYRILLIC_LETTER_RE = re.compile(r"[А-Яа-яЁё]")
LATIN_WORD_RE = re.compile(r"[A-Za-z][A-Za-z'-]*")

NAME_WORD = r"[A-Z][A-Za-z'-]+"
AUTHOR_PATTERNS = [
    re.compile(rf"\b{NAME_WORD}\s+[A-Z]\.\s*[A-Z]\.?\b"),
    re.compile(rf"\b[A-Z]\.\s*[A-Z]\.\s*{NAME_WORD}\b"),
    re.compile(rf"\b{NAME_WORD}\s+{NAME_WORD}\b"),
]
STRICT_AUTHOR_PATTERNS = AUTHOR_PATTERNS[:2]


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

        for line_no, line in enumerate(split_nonempty_lines(text), start=1):
            lines.append(
                WindowLine(
                    paragraph_index=paragraph_index,
                    line_no=line_no,
                    text=line,
                    paragraph_first_char_font_size=paragraph_first_char_font_size,
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


def is_latin_rich(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False
    return count_latin_letters(normalized) >= 8 and count_cyrillic_letters(normalized) == 0 and bool(LATIN_LETTER_RE.search(normalized))


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


def looks_like_author_line(text: str) -> bool:
    normalized = normalize_text(text)
    if not is_latin_rich(normalized):
        return False
    if is_keyword_line(normalized) or is_service_line(normalized):
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

    lowered = normalized.casefold()
    if any(cue in lowered for cue in AFFILIATION_CUES):
        return False
    if "@" in normalized:
        return False

    words = LATIN_WORD_RE.findall(normalized)
    if len(words) < 4 or len(words) > 30:
        return False
    if normalized.endswith("."):
        return False

    return True


def has_title_font_size(line: WindowLine) -> bool:
    if line.paragraph_first_char_font_size is None:
        return False
    return abs(line.paragraph_first_char_font_size - TITLE_FONT_SIZE) <= TITLE_FONT_TOLERANCE


def looks_like_post_title_author_line(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return False
    if not is_latin_rich(normalized):
        return False
    if is_keyword_line(normalized) or is_service_line(normalized):
        return False

    lowered = normalized.casefold()
    if any(cue in lowered for cue in AFFILIATION_CUES):
        return False
    if "@" in normalized or "http://" in lowered or "https://" in lowered or "doi" in lowered or "//" in normalized:
        return False
    if re.match(r"^\d+\.", normalized):
        return False

    matches = find_author_matches(normalized)
    if len(matches) >= 2:
        return True
    if len(matches) == 1 and ("," in normalized or "*" in normalized or any(char.isdigit() for char in normalized)):
        return True
    return False


def choose_from_keyword_window(lines: list[WindowLine]) -> tuple[WindowLine | None, WindowLine | None, WindowLine | None]:
    keyword_line: WindowLine | None = None

    for line in lines:
        if is_keyword_line(line.text):
            keyword_line = line
            break

    if keyword_line is None:
        return None, None, None

    keyword_index = lines.index(keyword_line)

    author_line: WindowLine | None = None
    for line in reversed(lines[:keyword_index]):
        if looks_like_author_line(line.text):
            author_line = line
            break

    title_anchor = lines.index(author_line) if author_line is not None else keyword_index

    title_line: WindowLine | None = None
    for line in reversed(lines[:title_anchor]):
        if looks_like_title_line(line.text):
            title_line = line
            break

    return keyword_line, title_line, author_line


def choose_by_forward_scan(lines: list[WindowLine]) -> tuple[WindowLine | None, WindowLine | None]:
    title_line: WindowLine | None = None

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

    return title_line, author_line


def choose_by_keyword_and_font14(
    lines: list[WindowLine],
) -> tuple[WindowLine | None, WindowLine | None, WindowLine | None]:
    keyword_line: WindowLine | None = None

    for line in lines:
        if is_keyword_line(line.text):
            keyword_line = line
            break

    if keyword_line is None:
        return None, None, None

    keyword_index = lines.index(keyword_line)
    title_line: WindowLine | None = None

    for line in reversed(lines[:keyword_index]):
        if has_title_font_size(line) and looks_like_title_line(line.text):
            title_line = line
            break

    if title_line is None:
        return keyword_line, None, None

    title_index = lines.index(title_line)
    author_line: WindowLine | None = None

    if title_index + 1 < len(lines):
        candidate = lines[title_index + 1]
        if candidate.paragraph_index == title_line.paragraph_index and looks_like_post_title_author_line(candidate.text):
            author_line = candidate

    if author_line is None:
        next_paragraph_index = title_line.paragraph_index + 1
        for candidate in lines[title_index + 1 :]:
            if candidate.paragraph_index < next_paragraph_index:
                continue
            if candidate.paragraph_index > next_paragraph_index:
                break
            if looks_like_post_title_author_line(candidate.text):
                author_line = candidate
                break

    return keyword_line, title_line, author_line


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
    input_path = Path(INPUT_WINDOWS_CSV)
    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    docx_path = Path(DOCX_PATH)
    if not docx_path.exists():
        raise FileNotFoundError(f"Файл не найден: {docx_path}")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    windows = read_article_windows(input_path)
    rows = build_debug_rows(docx_path, windows)

    csv_path = output_dir / OUTPUT_CSV
    txt_path = output_dir / OUTPUT_TXT

    write_debug_csv(csv_path, rows)
    write_text(txt_path, build_debug_text(rows))
    print_summary(rows)

    print("\nФайлы сохранены:")
    print(csv_path)
    print(txt_path)


if __name__ == "__main__":
    main()

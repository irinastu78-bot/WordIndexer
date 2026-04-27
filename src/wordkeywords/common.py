from __future__ import annotations

import csv
import re
import unicodedata
from pathlib import Path
from typing import Iterable

COMPACT_LOCANT_KEYWORD_RE = re.compile(
    r"^(?P<head>\d+)\((?P<inner>\d+)(?P<inside>[A-Za-z\u0400-\u04FF][^)]*)\)"
    r"(?P<locants>\d+(?:,\d+)+)(?P<tail>[A-Za-z\u0400-\u04FF].*)$"
)


def normalize_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\ufeff", "")
    text = text.replace("\r", "\n").replace("\x0b", "\n").replace("\x0c", "\n")
    text = text.replace("\xa0", " ")
    text = text.replace("\x1e", "-")
    text = text.replace("\x00", "")
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def clean_keyword_tail(text: str) -> str:
    text = normalize_text(text)
    text = re.split(r"\s*(?=(?:УДК|UDC)\b)", text, maxsplit=1)[0]
    text = text.strip(" \t\r\n-–—")
    text = re.sub(r"[;]+\s*$", "", text)
    return text


def clean_keyword(keyword: str) -> str:
    keyword = normalize_text(keyword)
    keyword = keyword.strip(" ,;.:")
    keyword = re.sub(r"\s*-\s*", "-", keyword)
    keyword = restore_compact_locant_hyphens(keyword)
    keyword = normalize_text(keyword)
    return keyword.strip(" ,;.:")


def restore_compact_locant_hyphens(keyword: str) -> str:
    match = COMPACT_LOCANT_KEYWORD_RE.match(keyword)
    if match is None:
        return keyword

    return (
        f"{match.group('head')}-"
        f"({match.group('inner')}-{match.group('inside')})-"
        f"{match.group('locants')}-{match.group('tail')}"
    )


def sort_keywords(items: Iterable[str]) -> list[str]:
    return sorted(items, key=lambda s: unicodedata.normalize("NFKC", s).casefold())


def read_csv_rows(path: Path, delimiter: str = ";") -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as file_obj:
        reader = csv.DictReader(file_obj, delimiter=delimiter)
        for row in reader:
            rows.append(row)
    return rows


def write_csv(path: Path, header: list[str], rows: Iterable[Iterable[object]], delimiter: str = ";") -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file_obj:
        writer = csv.writer(file_obj, delimiter=delimiter)
        writer.writerow(header)
        for row in rows:
            writer.writerow(list(row))


def write_text(path: Path, content: str) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file_obj:
        file_obj.write(content)


def parse_pages(pages_text: str) -> list[int]:
    pages_text = normalize_text(pages_text)
    if not pages_text:
        return []

    result: list[int] = []
    for part in pages_text.split(","):
        item = normalize_text(part)
        if item.isdigit():
            result.append(int(item))
    return result

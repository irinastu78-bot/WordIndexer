from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import clean_keyword, normalize_text, read_csv_rows, sort_keywords, write_csv, write_text


INPUT_DIR = ROOT_DIR / "output"
OUTPUT_DIR = ROOT_DIR / "output"
INPUT_CSV = INPUT_DIR / "fast_keyword_rows_v2.csv"

RU_CSV = "keyword_index_ru.csv"
EN_CSV = "keyword_index_en.csv"
RU_TXT = "keyword_index_ru.txt"
EN_TXT = "keyword_index_en.txt"
COMBINED_TXT = "keyword_indexes_for_word.txt"


def split_joined_keywords(text: str) -> list[str]:
    if not text:
        return []

    result: list[str] = []
    for part in text.split("|"):
        item = clean_keyword(part)
        if item:
            result.append(item)
    return result


def canonical_key(keyword: str, lang: str) -> str:
    cleaned = clean_keyword(keyword)
    if lang == "EN":
        return cleaned.casefold()
    return cleaned


def choose_display_value(existing: str | None, candidate: str, lang: str) -> str:
    candidate = clean_keyword(candidate)

    if not existing:
        return candidate

    if lang == "EN" and existing.islower() and not candidate.islower():
        return candidate

    return existing


def build_indexes(rows: list[dict[str, str]]) -> tuple[dict[str, set[int]], dict[str, set[int]]]:
    ru_pages_by_key: dict[str, set[int]] = defaultdict(set)
    en_pages_by_key: dict[str, set[int]] = defaultdict(set)

    ru_display_by_key: dict[str, str] = {}
    en_display_by_key: dict[str, str] = {}

    for row in rows:
        lang = normalize_text(row.get("lang", ""))
        page_raw = normalize_text(row.get("page", ""))
        keywords_joined = row.get("keywords_joined", "")

        if not page_raw.isdigit():
            continue

        page = int(page_raw)
        keywords = split_joined_keywords(keywords_joined)

        if lang == "RU":
            for keyword in keywords:
                key = canonical_key(keyword, "RU")
                if not key:
                    continue
                ru_pages_by_key[key].add(page)
                ru_display_by_key[key] = choose_display_value(ru_display_by_key.get(key), keyword, "RU")

        elif lang == "EN":
            for keyword in keywords:
                key = canonical_key(keyword, "EN")
                if not key:
                    continue
                en_pages_by_key[key].add(page)
                en_display_by_key[key] = choose_display_value(en_display_by_key.get(key), keyword, "EN")

    ru_index = {ru_display_by_key[key]: pages for key, pages in ru_pages_by_key.items()}
    en_index = {en_display_by_key[key]: pages for key, pages in en_pages_by_key.items()}

    return ru_index, en_index


def write_index_csv(path: Path, index: dict[str, set[int]]) -> None:
    rows = (
        [
            keyword,
            ", ".join(str(page) for page in sorted(index[keyword])),
            len(index[keyword]),
        ]
        for keyword in sort_keywords(index.keys())
    )
    write_csv(path, ["keyword", "pages", "pages_count"], rows)


def build_index_text(index: dict[str, set[int]]) -> str:
    lines: list[str] = []

    for keyword in sort_keywords(index.keys()):
        pages_sorted = sorted(index[keyword])
        lines.append(f"{keyword} {', '.join(str(p) for p in pages_sorted)}")

    return "\n".join(lines)


def build_combined_text(ru_index: dict[str, set[int]], en_index: dict[str, set[int]]) -> str:
    ru_text = build_index_text(ru_index)
    en_text = build_index_text(en_index)

    parts = [
        "ПРЕДМЕТНЫЙ УКАЗАТЕЛЬ",
        "",
        ru_text,
        "",
        "",
        "KEYWORD INDEX",
        "",
        en_text,
        "",
    ]
    return "\n".join(parts)


def print_summary(ru_index: dict[str, set[int]], en_index: dict[str, set[int]]) -> None:
    print("=" * 100)
    print(f"RU UNIQUE KEYWORDS:      {len(ru_index)}")
    print(f"EN UNIQUE KEYWORDS:      {len(en_index)}")
    print(f"TOTAL UNIQUE KEYWORDS:   {len(ru_index) + len(en_index)}")
    print("=" * 100)

    ru_preview = sort_keywords(ru_index.keys())[:10]
    en_preview = sort_keywords(en_index.keys())[:10]

    if ru_preview:
        print("\nПервые 10 RU keyword-фраз:")
        for kw in ru_preview:
            pages = ", ".join(str(p) for p in sorted(ru_index[kw]))
            print(f"- {kw} -> {pages}")

    if en_preview:
        print("\nПервые 10 EN keyword-фраз:")
        for kw in en_preview:
            pages = ", ".join(str(p) for p in sorted(en_index[kw]))
            print(f"- {kw} -> {pages}")


def main() -> None:
    input_path = Path(INPUT_CSV)
    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    rows = read_csv_rows(input_path)
    ru_index, en_index = build_indexes(rows)

    ru_csv_path = output_dir / RU_CSV
    en_csv_path = output_dir / EN_CSV
    ru_txt_path = output_dir / RU_TXT
    en_txt_path = output_dir / EN_TXT
    combined_txt_path = output_dir / COMBINED_TXT

    write_index_csv(ru_csv_path, ru_index)
    write_index_csv(en_csv_path, en_index)

    ru_text = build_index_text(ru_index)
    en_text = build_index_text(en_index)
    combined_text = build_combined_text(ru_index, en_index)

    write_text(ru_txt_path, ru_text)
    write_text(en_txt_path, en_text)
    write_text(combined_txt_path, combined_text)

    print_summary(ru_index, en_index)

    print("\nФайлы сохранены:")
    print(ru_csv_path)
    print(en_csv_path)
    print(ru_txt_path)
    print(en_txt_path)
    print(combined_txt_path)


if __name__ == "__main__":
    main()

from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
import re
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from wordkeywords.common import clean_keyword_tail, normalize_text, sort_keywords, write_csv

try:
    import win32com.client as win32
except ImportError:
    win32 = None


INPUT_DIR = ROOT_DIR / "input"
OUTPUT_DIR = ROOT_DIR / "output"
DEFAULT_DOCX_PATH = INPUT_DIR / "test1.docx"

RAW_ROWS_CSV = "fast_keyword_rows_v2.csv"
INDEX_CSV = "fast_keyword_index_v2.csv"


@dataclass
class KeywordHit:
    hit_no: int
    lang: str
    label: str
    page: int | None
    paragraph_text: str
    extracted_text: str
    keywords: list[str]


def build_tagged_name(filename: str, run_tag: str) -> str:
    return f"{run_tag}_{filename}" if run_tag else filename


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract keyword paragraphs from a .docx via Word COM and write CSV artifacts."
    )
    parser.add_argument(
        "--docx",
        default=str(DEFAULT_DOCX_PATH),
        help=f"Path to the input .docx file. Default: {DEFAULT_DOCX_PATH}",
    )
    parser.add_argument(
        "--run-tag",
        default="",
        help="Optional tag prefix for output files, e.g. 'run1'.",
    )
    return parser.parse_args(argv)


def split_keywords(text: str) -> list[str]:
    text = normalize_text(text)
    text = re.sub(r"[.;]\s*$", "", text)

    parts = re.split(r"(?<!\d),(?!\d)", text)

    result: list[str] = []
    for raw in parts:
        item = normalize_text(raw).strip(" ,;.")
        if item:
            result.append(item)
    return result


def parse_keyword_paragraph(paragraph_text: str, label: str) -> tuple[str, list[str]]:
    text = normalize_text(paragraph_text)
    if not text.lower().startswith(label.lower()):
        return "", []

    tail = text[len(label):].lstrip()
    tail = clean_keyword_tail(tail)
    keywords = split_keywords(tail)
    return tail, keywords


def get_page_number_from_range(rng) -> int | None:
    try:
        dup = rng.Duplicate
        dup.Collapse(1)  # wdCollapseStart
        return int(dup.Information(1))  # wdActiveEndAdjustedPageNumber
    except Exception:
        return None


def make_find(range_obj, label: str):
    find = range_obj.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Text = label
    find.Forward = True
    find.Wrap = 0  # wdFindStop
    find.Format = False
    find.MatchCase = False
    find.MatchWholeWord = False
    find.MatchWildcards = False
    find.MatchSoundsLike = False
    find.MatchAllWordForms = False
    return find


def extract_hits_by_label(doc, label: str, lang: str, hits: list[KeywordHit]) -> None:
    search_range = doc.Content.Duplicate
    search_range.Collapse(1)  # start of document

    while True:
        find = make_find(search_range, label)
        found = find.Execute()
        if not found:
            break

        found_range = search_range.Duplicate
        paragraph_range = found_range.Paragraphs(1).Range
        paragraph_text = normalize_text(paragraph_range.Text)

        extracted_text, keywords = parse_keyword_paragraph(paragraph_text, label)
        if keywords:
            page = get_page_number_from_range(paragraph_range)
            hits.append(
                KeywordHit(
                    hit_no=len(hits) + 1,
                    lang=lang,
                    label=label,
                    page=page,
                    paragraph_text=paragraph_text,
                    extracted_text=extracted_text,
                    keywords=keywords,
                )
            )

            if len(hits) % 20 == 0:
                print(f"HITS FOUND SO FAR: {len(hits)}")

        new_start = found_range.End
        if new_start >= doc.Content.End:
            break

        search_range = doc.Range(Start=new_start, End=doc.Content.End)


def deduplicate_hits(hits: list[KeywordHit]) -> list[KeywordHit]:
    seen: set[tuple[str, int | None, str]] = set()
    result: list[KeywordHit] = []

    for hit in hits:
        key = (hit.lang, hit.page, normalize_text(hit.paragraph_text).lower())
        if key in seen:
            continue
        seen.add(key)
        result.append(hit)

    for i, hit in enumerate(result, start=1):
        hit.hit_no = i

    return result


def extract_keyword_hits_with_word(docx_path: str) -> list[KeywordHit]:
    if win32 is None:
        raise RuntimeError("Не установлен pywin32. Установите: pip install pywin32")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    doc = None
    hits: list[KeywordHit] = []

    try:
        doc = word.Documents.Open(str(Path(docx_path).resolve()), ReadOnly=True)

        extract_hits_by_label(doc, "Ключевые слова:", "RU", hits)
        extract_hits_by_label(doc, "Keywords:", "EN", hits)
        extract_hits_by_label(doc, "Ketwords:", "EN", hits)
        extract_hits_by_label(doc, "Key words:", "EN", hits)

        return deduplicate_hits(hits)

    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


def build_index(hits: list[KeywordHit]) -> dict[str, set[int]]:
    index: dict[str, set[int]] = defaultdict(set)

    for hit in hits:
        if hit.page is None:
            continue
        for keyword in hit.keywords:
            index[keyword].add(hit.page)

    return index


def write_rows_csv(path: Path, hits: list[KeywordHit]) -> None:
    rows = (
        [
            hit.hit_no,
            hit.lang,
            hit.label,
            hit.page if hit.page is not None else "",
            hit.paragraph_text,
            hit.extracted_text,
            " | ".join(hit.keywords),
        ]
        for hit in hits
    )
    write_csv(
        path,
        [
            "hit_no",
            "lang",
            "label",
            "page",
            "paragraph_text",
            "extracted_text",
            "keywords_joined",
        ],
        rows,
    )


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


def print_summary(hits: list[KeywordHit], index: dict[str, set[int]]) -> None:
    ru_hits = sum(1 for h in hits if h.lang == "RU")
    en_hits = sum(1 for h in hits if h.lang == "EN")
    no_page = sum(1 for h in hits if h.page is None)

    print("=" * 100)
    print(f"TOTAL HITS:              {len(hits)}")
    print(f"RU HITS:                 {ru_hits}")
    print(f"EN HITS:                 {en_hits}")
    print(f"HITS WITHOUT PAGE:       {no_page}")
    print(f"UNIQUE KEYWORDS:         {len(index)}")
    print("=" * 100)

    preview = sort_keywords(index.keys())[:20]
    if preview:
        print("\nПервые 20 keyword-фраз:")
        for kw in preview:
            pages = ", ".join(str(p) for p in sorted(index[kw]))
            print(f"- {kw} -> {pages}")


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)

    docx_path = Path(args.docx)
    if not docx_path.exists():
        raise FileNotFoundError(f"Файл не найден: {docx_path}")

    if docx_path.suffix.lower() != ".docx":
        raise ValueError(f"Expected a .docx file: {docx_path}")

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    rows_csv_path = output_dir / build_tagged_name(RAW_ROWS_CSV, args.run_tag)
    index_csv_path = output_dir / build_tagged_name(INDEX_CSV, args.run_tag)

    print(f"INPUT DOCX: {docx_path}")
    print(f"RUN TAG:    {args.run_tag or '(none)'}")

    hits = extract_keyword_hits_with_word(str(docx_path))
    index = build_index(hits)

    write_rows_csv(rows_csv_path, hits)
    write_index_csv(index_csv_path, index)
    print_summary(hits, index)

    print("\nФайлы сохранены:")
    print(rows_csv_path)
    print(index_csv_path)


if __name__ == "__main__":
    main()

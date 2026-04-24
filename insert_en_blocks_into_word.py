from __future__ import annotations

import argparse
from pathlib import Path

from scripts.insert_keyword_indexes_into_word import (
    insert_heading,
    insert_page_break,
    insert_plain_paragraph,
)

try:
    import win32com.client as win32
except ImportError:
    win32 = None


OUTPUT_DIR = Path("output")
AUTHOR_PROGRESS_STEP = 25


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Insert EN final blocks into a separate Word output document."
    )
    parser.add_argument(
        "--docx",
        dest="docx_path",
        type=Path,
        required=True,
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
        return OUTPUT_DIR / "final_with_toc_author_en.docx"
    return OUTPUT_DIR / f"{clean_run_tag}_final_with_toc_author_en.docx"


def resolve_author_index_txt_path(run_tag: str) -> Path:
    primary_path = resolve_output_path("draft_author_index_en.txt", run_tag)
    fallback_path = resolve_output_path("draft_author_index_en.txt", "")

    if primary_path.exists() and fallback_path.exists() and primary_path.resolve() != fallback_path.resolve():
        if fallback_path.stat().st_mtime >= primary_path.stat().st_mtime:
            print(f"EN author index source fallback selected (newer file): {fallback_path.resolve()}")
            return fallback_path

    if primary_path.exists():
        return primary_path
    if run_tag.strip() and fallback_path.exists():
        return fallback_path
    return primary_path


def build_numbered_docx_path(path: Path, index: int) -> Path:
    suffix = path.suffix or ".docx"
    return path.with_name(f"{path.stem}_{index}{suffix}")


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


def get_com_error_details(error) -> tuple[int | None, list[str]]:
    code = getattr(error, "hresult", None)
    details: list[str] = []

    excepinfo = getattr(error, "excepinfo", None)
    if excepinfo:
        if isinstance(excepinfo, tuple) and excepinfo:
            scode = excepinfo[-1]
            if isinstance(scode, int):
                code = scode
        details.extend(iter_error_text_parts(excepinfo))

    args = getattr(error, "args", None)
    if args:
        if not isinstance(args, tuple):
            args = (args,)
        if args:
            first = args[0]
            if isinstance(first, int):
                code = first
        for item in args:
            details.extend(iter_error_text_parts(item))

    seen: set[str] = set()
    unique_details: list[str] = []
    for item in details:
        key = item.casefold()
        if key not in seen:
            seen.add(key)
            unique_details.append(item)

    return code, unique_details


def is_open_document_name_conflict(error) -> bool:
    code, details = get_com_error_details(error)
    if code in (-2146823683, -2147352567):
        detail_text = " ".join(details).casefold()
        if (
            "same name" in detail_text
            or "already open" in detail_text
            or "open document with the same name" in detail_text
        ):
            return True

    text = " ".join(details).casefold()
    if (
        "same name" in text
        or "already open" in text
        or "open document with the same name" in text
    ):
        return True

    return False


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


def read_nonempty_lines(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    lines: list[str] = []
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = " ".join(raw_line.split()).strip()
        if line:
            lines.append(line)
    return lines


def main() -> int:
    args = parse_args()
    if win32 is None:
        raise RuntimeError("pywin32 is not installed. Install it with: pip install pywin32")

    run_tag = args.run_tag.strip()
    source_path = Path(args.docx_path).resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"File not found: {source_path}")

    author_txt_path = resolve_author_index_txt_path(run_tag)
    output_docx_path = resolve_output_docx_path(run_tag)
    en_toc_csv_path = resolve_output_path("draft_toc_en.csv", run_tag)

    author_lines = read_nonempty_lines(author_txt_path)
    output_docx_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"source docx path: {source_path}")
    print(f"EN author index source: {author_txt_path.resolve()}")
    if en_toc_csv_path.exists():
        print(f"EN TOC source detected but not inserted in this step: {en_toc_csv_path.resolve()}")
    else:
        print("EN TOC source not found; skipping EN TOC block.")
    print(f"EN author lines: {len(author_lines)}")
    print(f"output docx path: {output_docx_path.resolve()}")
    print("opening Word...")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    document = None
    try:
        document = word.Documents.Open(str(source_path))

        print("inserting EN author heading...")
        insert_page_break(document)
        insert_heading(document, "Author Index", level=1)

        print("inserting EN author lines...")
        for index, line in enumerate(author_lines, start=1):
            insert_plain_paragraph(document, line)
            if index % AUTHOR_PROGRESS_STEP == 0 or index == len(author_lines):
                print(f"EN author lines inserted: {index}/{len(author_lines)}")

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

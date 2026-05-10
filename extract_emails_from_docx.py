from pathlib import Path
import re
import zipfile
import html
import csv

DOCX_PATH = Path("output/test7_final_ordered.docx")
OUT_TXT = Path("output/test7_emails.txt")
OUT_CSV = Path("output/test7_emails.csv")

EMAIL_RE = re.compile(
    r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b",
    re.IGNORECASE,
)

def extract_docx_text(docx_path: Path) -> str:
    parts = []

    with zipfile.ZipFile(docx_path) as z:
        for name in z.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                xml = z.read(name).decode("utf-8", errors="ignore")
                xml = re.sub(r"<[^>]+>", " ", xml)
                xml = html.unescape(xml)
                parts.append(xml)

    return "\n".join(parts)

def main():
    if not DOCX_PATH.exists():
        raise FileNotFoundError(DOCX_PATH)

    text = extract_docx_text(DOCX_PATH)
    emails = sorted(set(m.group(0) for m in EMAIL_RE.finditer(text)))

    OUT_TXT.parent.mkdir(parents=True, exist_ok=True)

    OUT_TXT.write_text("\n".join(emails), encoding="utf-8")

    with OUT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["email"])
        for email in emails:
            writer.writerow([email])

    print(f"Emails found: {len(emails)}")
    print(f"TXT: {OUT_TXT}")
    print(f"CSV: {OUT_CSV}")

if __name__ == "__main__":
    main()
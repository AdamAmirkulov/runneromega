"""
Splits Bereke Bank prt-*.pdf files (each containing a Досудебная претензия
followed by an Уведомление об уступки права требования) into two separate
PDFs per source file, named:

    <Название документа>, <ФИО>, <ИИН>.pdf

Source PDFs are left untouched; output goes into an "output" subfolder next
to the source files. Any file the script can't confidently parse is skipped
and logged to failed.log for manual handling.
"""

import argparse
import glob
import os
import re
import sys

import pdfplumber
from pypdf import PdfReader, PdfWriter

NAME_CHARS = r"А-ЯЁӘҒҚҢӨҰҮҺІ"
PAT_HEADER = re.compile(
    rf"кому[;:]\s*([{NAME_CHARS}][{NAME_CHARS}\s\-]+?)\s*ИИН[:\s]*(\d{{12}})",
    re.IGNORECASE,
)
PAT_NOTICE = re.compile(
    rf"между\s+Вами\s+([{NAME_CHARS}][{NAME_CHARS}\s\-]+?)\s*ИИН[:\s]*(\d{{12}})",
    re.IGNORECASE,
)

DOC_PRETENZIA = "Досудебная претензия"
DOC_UVEDOMLENIE = "Уведомление об уступки права"

INVALID_CHARS = re.compile(r'[\\/:*?"<>|]')


def clean_name(raw: str) -> str:
    raw = re.sub(r"\s+", " ", raw).strip()
    raw = re.sub(r"\s*-\s*", "-", raw)  # collapse "УСМАНОВА- АБДЫКАСЫМОВА" -> "УСМАНОВА-АБДЫКАСЫМОВА"
    return raw


def left_column_text(page) -> str:
    """Reconstruct the Russian (left) column, ignoring the Kazakh column that
    pdfplumber otherwise interleaves word-by-word into the same line."""
    mid = page.width / 2
    words = [w for w in page.extract_words() if w["x0"] < mid]
    lines = {}
    for w in words:
        key = round(w["top"] / 3)
        lines.setdefault(key, []).append(w)
    out = []
    for key in sorted(lines):
        row = sorted(lines[key], key=lambda x: x["x0"])
        out.append(" ".join(x["text"] for x in row))
    return " ".join(out)


def find_notice_page(pdf) -> int | None:
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        if "Уведомление об уступ" in text:
            return i
    return None


def extract_name_iin(pdf, notice_idx: int | None):
    p0 = left_column_text(pdf.pages[0])
    m = PAT_HEADER.search(p0)
    if not m and notice_idx is not None:
        pn = left_column_text(pdf.pages[notice_idx])
        m = PAT_NOTICE.search(pn)
    if not m:
        return None, None
    return clean_name(m.group(1)), m.group(2)


def safe_filename(doc_title: str, name: str, iin: str) -> str:
    base = f"{doc_title}, {name}, {iin}.pdf"
    return INVALID_CHARS.sub("_", base)


def unique_path(directory: str, filename: str) -> str:
    path = os.path.join(directory, filename)
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(filename)
    n = 2
    while True:
        candidate = os.path.join(directory, f"{stem} ({n}){ext}")
        if not os.path.exists(candidate):
            return candidate
        n += 1


def write_pages(src_path: str, page_indices, out_path: str):
    reader = PdfReader(src_path)
    writer = PdfWriter()
    for i in page_indices:
        writer.add_page(reader.pages[i])
    with open(out_path, "wb") as f:
        writer.write(f)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--source_folder",
        default=None,
        help="Folder containing prt-*.pdf files (as passed by the web app)",
    )
    parser.add_argument(
        "--folder",
        default=None,
        help="Alternative name for the source folder",
    )
    parser.add_argument(
        "--workdir",
        default=None,
        help="Job working directory passed by the runner (unused, accepted for compatibility)",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="Output folder (default: <folder>\\output)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only report what would be done, do not write any files",
    )
    args, _ = parser.parse_known_args()

    folder = (
        args.source_folder
        or args.folder
        or r"C:\Обмен\work folder\Уведомления об уступки права требования\Bereke Bank"
    )
    out_dir = args.out or os.path.join(folder, "output")
    if not args.dry_run:
        os.makedirs(out_dir, exist_ok=True)

    files = sorted(glob.glob(os.path.join(folder, "prt-*.pdf")))
    print(f"Found {len(files)} files in {folder}")

    log_path = os.path.join(folder, "split_rename_log.txt")
    ok_count = 0
    fail_count = 0

    with open(log_path, "w", encoding="utf-8") as log:
        for path in files:
            fname = os.path.basename(path)
            try:
                with pdfplumber.open(path) as pdf:
                    notice_idx = find_notice_page(pdf)
                    if notice_idx is None or notice_idx == 0:
                        log.write(f"{fname}\tSKIP\tcould not locate notice section\n")
                        fail_count += 1
                        continue
                    name, iin = extract_name_iin(pdf, notice_idx)
                    n_pages = len(pdf.pages)
            except Exception as e:
                log.write(f"{fname}\tERROR\t{e}\n")
                fail_count += 1
                continue

            if not name or not iin:
                log.write(f"{fname}\tSKIP\tcould not extract FIO/IIN\n")
                fail_count += 1
                continue

            pretenzia_range = range(0, notice_idx)
            notice_range = range(notice_idx, n_pages)

            pretenzia_name = safe_filename(DOC_PRETENZIA, name, iin)
            notice_name = safe_filename(DOC_UVEDOMLENIE, name, iin)

            if args.dry_run:
                log.write(
                    f"{fname}\tOK\t{name}\t{iin}\t"
                    f"pretenzia_pages={list(pretenzia_range)}\t"
                    f"notice_pages={list(notice_range)}\n"
                )
                ok_count += 1
                continue

            try:
                pretenzia_out = unique_path(out_dir, pretenzia_name)
                write_pages(path, pretenzia_range, pretenzia_out)

                notice_out = unique_path(out_dir, notice_name)
                write_pages(path, notice_range, notice_out)

                log.write(f"{fname}\tOK\t{name}\t{iin}\n")
                ok_count += 1
            except Exception as e:
                log.write(f"{fname}\tERROR\twrite failed: {e}\n")
                fail_count += 1

    print(f"Done. OK={ok_count} FAILED={fail_count}")
    print(f"Log written to: {log_path}")
    if fail_count:
        print("Review the log for skipped/failed files and handle them manually.")


if __name__ == "__main__":
    sys.exit(main())

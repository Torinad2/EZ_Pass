#!/usr/bin/env python3
"""
EZ-Pass (NY/PANYNJ style) statement PDF -> Excel (.xlsx)

Usage:
  python ezpass_pdf_to_excel.py input.pdf output.xlsx
  python ezpass_pdf_to_excel.py input_folder/ output.xlsx

Notes:
- Uses pdfplumber for text extraction (works best for "text" PDFs, not scanned images).
- The parser is tolerant of "Service Fee" / "Prepaid Toll Payment" lines that don't have tag/agency fields.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd
import pdfplumber

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2}$")
TIME_RE = re.compile(r"^\d{2}:\d{2}$")
MONEY_RE = re.compile(r"^-?\$[\d,]+\.\d{2}$")
TAG_RE = re.compile(r"^\d{8,}$")  # EZ-Pass tag/plate often long digits


def money_to_float(s: str) -> Optional[float]:
    if s is None:
        return None
    s = s.strip()
    if not s:
        return None
    # Expected like "$1.23" or "-$6.26"
    neg = s.startswith("-")
    s = s.replace("$", "").replace(",", "").replace("-", "")
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None


def mmddyy_to_date(s: str) -> Optional[pd.Timestamp]:
    if not s:
        return None
    try:
        return pd.to_datetime(s, format="%m/%d/%y")
    except Exception:
        return None


def parse_statement_metadata(full_text: str) -> dict:
    def find(pattern: str) -> Optional[re.Match]:
        return re.search(pattern, full_text, flags=re.MULTILINE)

    stmt_date = find(r"Statement Date:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})")
    acct = find(r"Account Number:\s*(\d+)")
    agency = find(r"Agency:\s*(.+)")
    activity = find(
        r"Activity For:\s*([0-9]{2}/[0-9]{2}/[0-9]{2})\s*-\s*([0-9]{2}/[0-9]{2}/[0-9]{2})"
    )

    beginning = find(r"Beginning Balance\s*([-\$0-9\.,]+)")
    ending = find(r"Ending Balance\s*([-\$0-9\.,]+)")
    tolls_fees = find(r"Tolls, Fees and Parking.*?\s*([-\$0-9\.,]+)")
    payments = find(r"Payments/Adjustments\s*([-\$0-9\.,]+)")

    return {
        "statement_date": stmt_date.group(1) if stmt_date else None,
        "account_number": acct.group(1) if acct else None,
        "agency": agency.group(1).strip() if agency else None,
        "activity_start": activity.group(1) if activity else None,
        "activity_end": activity.group(2) if activity else None,
        "beginning_balance": beginning.group(1) if beginning else None,
        "tolls_fees": tolls_fees.group(1) if tolls_fees else None,
        "payments_adjustments": payments.group(1) if payments else None,
        "ending_balance": ending.group(1) if ending else None,
    }


def iter_transaction_lines(page_text: str) -> Iterable[str]:
    """
    Pull only the transaction section from the page that contains it.
    """
    start = page_text.find("POSTING")
    if start == -1:
        return []
    end = page_text.find("PREPAID TOLL BALANCE")
    if end == -1:
        end = len(page_text)

    block = page_text[start:end]
    lines = [ln.strip() for ln in block.splitlines() if ln.strip()]

    # Skip header lines; yield only lines beginning with a date
    for ln in lines:
        if DATE_RE.match(ln.split()[0]):
            yield ln


def parse_transaction_line(line: str) -> Optional[dict]:
    toks = line.split()
    if len(toks) < 4:
        return None
    if not (DATE_RE.match(toks[0]) and DATE_RE.match(toks[1])):
        return None

    posting_date, transaction_date = toks[0], toks[1]
    amount, balance = toks[-2], toks[-1]

    # Non-toll rows (fees, payments): date date description $amount $balance
    if not (len(toks) >= 6 and TAG_RE.match(toks[2])):
        desc = " ".join(toks[2:-2])
        return {
            "posting_date": posting_date,
            "transaction_date": transaction_date,
            "tag_or_plate": None,
            "agency": None,
            "plaza": None,
            "entry_date": None,
            "entry_time": None,
            "exit_plaza": None,
            "exit_date": None,
            "exit_time": None,
            "plan": None,
            "vehicle_class": None,
            "amount": amount,
            "balance": balance,
            "description": desc,
        }

    # Toll row
    tag = toks[2]
    agency = toks[3]
    plaza = toks[4]
    plan = toks[-4]
    vehicle_class = toks[-3]
    middle = toks[5:-4]

    entry_date = entry_time = exit_plaza = exit_date = exit_time = None

    if middle:
        # Sometimes middle begins with a date; sometimes it begins with an entry plaza.
        if DATE_RE.match(middle[0]):
            entry_date = middle[0]
            if len(middle) > 1 and TIME_RE.match(middle[1]):
                entry_time = middle[1]
            rest = middle[2:]
        else:
            # Treat the first token as an entry plaza, and shift the rest.
            plaza = middle[0]
            rest = middle[1:]
            if rest and DATE_RE.match(rest[0]):
                entry_date = rest[0]
                if len(rest) > 1 and TIME_RE.match(rest[1]):
                    entry_time = rest[1]
                rest = rest[2:]

        # Remaining could be exit plaza/date/time (if present)
        if rest:
            exit_plaza = rest[0]
            if len(rest) > 1 and DATE_RE.match(rest[1]):
                exit_date = rest[1]
            if len(rest) > 2 and TIME_RE.match(rest[2]):
                exit_time = rest[2]

    return {
        "posting_date": posting_date,
        "transaction_date": transaction_date,
        "tag_or_plate": tag,
        "agency": agency,
        "plaza": plaza,
        "entry_date": entry_date,
        "entry_time": entry_time,
        "exit_plaza": exit_plaza,
        "exit_date": exit_date,
        "exit_time": exit_time,
        "plan": plan,
        "vehicle_class": vehicle_class,
        "amount": amount,
        "balance": balance,
        "description": None,
    }


def parse_pdf(pdf_path: Path) -> tuple[dict, pd.DataFrame]:
    with pdfplumber.open(str(pdf_path)) as pdf:
        page_texts = [(page.extract_text() or "") for page in pdf.pages]

    full_text = "\n".join(page_texts)
    metadata = parse_statement_metadata(full_text)

    all_rows = []
    for page_text in page_texts:
        for line in iter_transaction_lines(page_text):
            row = parse_transaction_line(line)
            if row:
                row["source_pdf"] = pdf_path.name
                all_rows.append(row)

    df = pd.DataFrame(all_rows)

    # Clean up types for Excel
    if not df.empty:
        for col in ["posting_date", "transaction_date", "entry_date", "exit_date"]:
            df[col] = df[col].apply(mmddyy_to_date)
        for col in ["amount", "balance"]:
            df[col + "_num"] = df[col].apply(money_to_float)

    return metadata, df


def parse_many(inputs: list[Path]) -> tuple[pd.DataFrame, pd.DataFrame]:
    meta_rows = []
    tx_frames = []

    for p in inputs:
        meta, tx = parse_pdf(p)
        meta_rows.append({"source_pdf": p.name, **meta})
        tx_frames.append(tx)

    meta_df = pd.DataFrame(meta_rows)
    tx_df = pd.concat(tx_frames, ignore_index=True) if tx_frames else pd.DataFrame()
    return meta_df, tx_df


def collect_inputs(path: Path) -> list[Path]:
    if path.is_file():
        return [path]
    if path.is_dir():
        pdfs = sorted(path.glob("*.pdf"))
        if not pdfs:
            raise FileNotFoundError(f"No PDF files found in folder: {path}")
        return pdfs
    raise FileNotFoundError(f"Not found: {path}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("input", type=str, help="PDF file or folder of PDFs")
    ap.add_argument("output", type=str, help="Output .xlsx file")
    args = ap.parse_args()

    in_path = Path(args.input).expanduser().resolve()
    out_path = Path(args.output).expanduser().resolve()

    inputs = collect_inputs(in_path)
    meta_df, tx_df = parse_many(inputs)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        meta_df.to_excel(writer, sheet_name="Statements", index=False)
        tx_df.to_excel(writer, sheet_name="Transactions", index=False)

        # Make sheets readable
        for sheet_name in ["Statements", "Transactions"]:
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col[:2000]:
                    if cell.value is None:
                        continue
                    max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 45)

    print(f"Saved Excel: {out_path}")


if __name__ == "__main__":
    main()

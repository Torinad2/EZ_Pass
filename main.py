#!/usr/bin/env python3

"""
EZ-Pass PDF Statement Parser
----------------------------
This program extracts transaction data from EZ-Pass PDF statements
and exports the data into an Excel file.

The program supports:
- One PDF file OR a folder with multiple PDFs
- Clean Excel output
- Dates preserved as MM/DD/YY strings
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd
import pdfplumber

# -------------------------------------------------
# Main program entry point
# -------------------------------------------------
def main() -> None:
    # Command-line arguments
    ap = argparse.ArgumentParser()
    ap.add_argument("input", type=str, help="PDF file or folder with PDFs")
    ap.add_argument("output", type=str, help="Output Excel file (.xlsx)")
    args = ap.parse_args()

    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve()

    # Collect PDF files
    pdf_files = collect_inputs(input_path)

    # Parse PDFs into Transactions table
    tx_df = parse_many(pdf_files)

    # Write Excel file
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        tx_df.to_excel(writer, sheet_name="Transactions", index=False)

        ws = writer.book["Transactions"]
        ws.freeze_panes = "A2"

        # Auto-size columns
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col[:2000]:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 45)

    print(f"Saved Excel: {output_path}")

# -------------------------------------------------
# Regular expressions used for parsing
# -------------------------------------------------
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2}$")
TIME_RE = re.compile(r"^\d{2}:\d{2}$")
MONEY_RE = re.compile(r"^-?\$[\d,]+\.\d{2}$")
TAG_RE = re.compile(r"^\d{8,}$")  # EZ-Pass tag / plate number


# -------------------------------------------------
# Convert money string like "$1.74" to float 1.74
# -------------------------------------------------
def money_to_float(value: str) -> Optional[float]:
    if not value:
        return None

    value = value.strip()
    negative = value.startswith("-")

    value = value.replace("$", "").replace(",", "").replace("-", "")

    try:
        amount = float(value)
        return -amount if negative else amount
    except ValueError:
        return None


# -------------------------------------------------
# Extract only the transaction lines from a page
# -------------------------------------------------
def iter_transaction_lines(page_text: str) -> Iterable[str]:
    """
    Return all lines that look like transaction rows.

    Supports:
    - New format:   LANE_TXN_ID POSTED_DATE ...
      Example:      31420710413 04/06/25 ...
    - Fee format:   POSTED_DATE Description AMT BAL
      Example:      04/03/25 Monthly Service Fee -$1.00 -$19.91
    - Old format:   POSTING_DATE TXN_DATE ...
      Example:      12/11/24 12/10/24 ...
    """
    for line in (ln.strip() for ln in page_text.splitlines()):
        if not line:
            continue

        toks = line.split()
        if len(toks) < 4:
            continue

        # New format with lane id: digits + date
        if toks[0].isdigit() and len(toks) > 1 and DATE_RE.match(toks[1]):
            yield line
            continue

        # Old/fee formats: date at the beginning
        if DATE_RE.match(toks[0]):
            yield line


# -------------------------------------------------
# Parse one transaction line into structured data
# -------------------------------------------------
def parse_transaction_line(line: str) -> Optional[dict]:
    """
    Parses a single statement line into a normalized transaction row.

    Supports two layouts:
    1) Older layout:
         POSTING_DATE TXN_DATE TAG AGENCY PLAZA ... PLAN CL AMT BALANCE
    2) Newer layout with LANE TXN ID:
         LANE_TXN_ID POSTED_DATE TAG/PLATE/DESCRIPTION AGENCY PLAZA ENTRY_DATE ENTRY_TIME ... PLAN CL AMT BALANCE

    Also supports fee/payment rows that may have only:
         POSTED_DATE Description AMT BALANCE
    """
    tokens = line.split()
    if len(tokens) < 4:
        return None

    # Helper: create a row with the full schema (keep missing columns as None)
    def empty_row() -> dict:
        return {
            "lane_txn_id": None,
            "posting_date": None,          # posted date in statement
            "transaction_date": None,      # old format had 2 dates (posting + transaction)
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
            "amount": None,
            "balance": None,
            "description": None,
        }

    # Must end with AMT and BALANCE (money strings like $.. or -$..)
    # We won't strictly validate with MONEY_RE because some statements can vary,
    # but we assume last 2 tokens are amount + balance in these EZ-Pass PDFs.
    amount = tokens[-2]
    balance = tokens[-1]

    # ------------------------------------------------------------
    # FORMAT B (new): starts with lane txn id (digits), then a date
    # Example from your new PDF:
    # 31420710413 04/06/25 00504721314 MTAB&T BWB 04/04/25 04:27 STANDARD 31 -$6.94 $73.15
    # ------------------------------------------------------------
    if tokens[0].isdigit() and DATE_RE.match(tokens[1]):
        row = empty_row()
        row["lane_txn_id"] = tokens[0]
        row["posting_date"] = tokens[1]

        # Fee/payment rows usually do NOT start with lane id, so here we treat as toll-like row.
        # Minimum expected: lane_id, posted_date, tag/plate, agency, plaza, entry_date, entry_time, plan, class, amt, balance
        # But sometimes some columns can be missing — we keep safe parsing.
        if len(tokens) < 9:
            # Not enough tokens to be meaningful
            row["amount"] = amount
            row["balance"] = balance
            row["description"] = " ".join(tokens[2:-2]) if len(tokens) > 4 else None
            return row

        row["tag_or_plate"] = tokens[2]
        row["agency"] = tokens[3] if len(tokens) > 3 else None
        row["plaza"] = tokens[4] if len(tokens) > 4 else None

        # Everything between plaza and the tail (PLAN CL AMT BALANCE) is the “middle”
        # Tail is usually: PLAN, CL, AMT, BALANCE (4 tokens)
        # So plan=tokens[-4], class=tokens[-3]
        if len(tokens) >= 6:
            row["plan"] = tokens[-4] if len(tokens) >= 6 else None
            row["vehicle_class"] = tokens[-3] if len(tokens) >= 5 else None

        middle = tokens[5:-4]  # after plaza up to before plan/class/amt/bal

        # Most common in your new PDF: middle = [ENTRY_DATE, ENTRY_TIME] (no exit fields)
        if middle:
            if DATE_RE.match(middle[0]):
                row["entry_date"] = middle[0]
                if len(middle) > 1 and TIME_RE.match(middle[1]):
                    row["entry_time"] = middle[1]
                rest = middle[2:]
            else:
                # If some statements insert extra plaza token, handle gracefully
                rest = middle

            # If exit info exists (rare in your sample, but possible):
            # rest could be: EXIT_PLAZA EXIT_DATE EXIT_TIME
            if rest:
                row["exit_plaza"] = rest[0]
                if len(rest) > 1 and DATE_RE.match(rest[1]):
                    row["exit_date"] = rest[1]
                if len(rest) > 2 and TIME_RE.match(rest[2]):
                    row["exit_time"] = rest[2]

        row["amount"] = amount
        row["balance"] = balance
        return row

    # ------------------------------------------------------------
    # FORMAT A (old): starts with a posting date and then transaction date
    # Example (older file):
    # 12/11/24 12/10/24 00504419585 ... $1.74 $16.74
    # ------------------------------------------------------------
    if DATE_RE.match(tokens[0]) and len(tokens) > 1 and DATE_RE.match(tokens[1]):
        row = empty_row()
        row["posting_date"] = tokens[0]
        row["transaction_date"] = tokens[1]

        # If 3rd token looks like a tag/plate -> toll row
        # Otherwise it’s fee/payment row
        row["tag_or_plate"] = tokens[2] if len(tokens) > 2 else None

        # Heuristic: toll rows usually have many tokens; fee rows often are short
        if len(tokens) <= 6:
            row["description"] = " ".join(tokens[2:-2])
            row["amount"] = amount
            row["balance"] = balance
            return row

        # Old layout fields (best-effort)
        row["agency"] = tokens[3] if len(tokens) > 3 else None
        row["plaza"] = tokens[4] if len(tokens) > 4 else None

        row["plan"] = tokens[-4] if len(tokens) >= 6 else None
        row["vehicle_class"] = tokens[-3] if len(tokens) >= 5 else None

        middle = tokens[5:-4]

        # Similar logic: parse entry/exit if present
        if middle:
            if DATE_RE.match(middle[0]):
                row["entry_date"] = middle[0]
                if len(middle) > 1 and TIME_RE.match(middle[1]):
                    row["entry_time"] = middle[1]
                rest = middle[2:]
            else:
                rest = middle

            if rest:
                row["exit_plaza"] = rest[0]
                if len(rest) > 1 and DATE_RE.match(rest[1]):
                    row["exit_date"] = rest[1]
                if len(rest) > 2 and TIME_RE.match(rest[2]):
                    row["exit_time"] = rest[2]

        row["amount"] = amount
        row["balance"] = balance
        return row

    # ------------------------------------------------------------
    # Fee/payment rows (new PDF style): only one date at start
    # Example from your new PDF:
    # 04/03/25 Monthly Service Fee -$1.00 -$19.91
    # 04/05/25 Prepaid Toll Payment $100.00 $80.09
    # ------------------------------------------------------------
    if DATE_RE.match(tokens[0]):
        row = empty_row()
        row["posting_date"] = tokens[0]
        row["description"] = " ".join(tokens[1:-2]) if len(tokens) > 3 else None
        row["amount"] = amount
        row["balance"] = balance
        return row

    return None



# -------------------------------------------------
# Parse a single PDF file
# -------------------------------------------------
def parse_pdf(pdf_path: Path) -> pd.DataFrame:
    """
    Opens a PDF file and extracts all transaction rows.
    Returns a pandas DataFrame.
    """
    rows = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in iter_transaction_lines(text):
                row = parse_transaction_line(line)
                if row:
                    rows.append(row)

    df = pd.DataFrame(rows)

    # Keep date columns as strings (MM/DD/YY)
    for col in ["posting_date", "transaction_date", "entry_date", "exit_date"]:
        if col in df.columns:
            df[col] = df[col].where(df[col].notna(), None)

    # Create numeric columns for calculations
    if not df.empty:
        df["amount_num"] = df["amount"].apply(money_to_float)
        df["balance_num"] = df["balance"].apply(money_to_float)

    return df


# -------------------------------------------------
# Parse one file or many files
# -------------------------------------------------
def parse_many(inputs: list[Path]) -> pd.DataFrame:
    """
    Parses multiple PDFs and combines all transactions
    into a single DataFrame.
    """
    frames = []

    for pdf in inputs:
        frames.append(parse_pdf(pdf))

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


# -------------------------------------------------
# Collect PDF files from input
# -------------------------------------------------
def collect_inputs(path: Path) -> list[Path]:
    """
    Accepts either:
    - One PDF file
    - A folder containing PDF files
    """
    if path.is_file():
        return [path]

    if path.is_dir():
        pdfs = sorted(path.glob("*.pdf"))
        if not pdfs:
            raise FileNotFoundError("No PDF files found in folder.")
        return pdfs

    raise FileNotFoundError("Input path not found.")

# -------------------------------------------------
# Run program
# -------------------------------------------------
if __name__ == "__main__":
    main()


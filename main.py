#!/usr/bin/env python3

#!/usr/bin/env python3

"""
/**************************************************************************
 * EZ-Pass PDF Statement Parser                                           *
 * ---------------------------------------------------------------------- *
 * Purpose                                                                *
 *   Extract transaction data from EZ-Pass PDF statements and export it   *
 *   into a clean Excel (.xlsx) file.                                     *
 *                                                                        *
 * What it supports                                                       *
 *   1) Input can be:                                                     *
 *        - One PDF file                                                  *
 *        - A folder containing multiple PDFs                             *
 *   2) Handles multiple statement layouts (old + new)                    *
 *        - Old layout with POSTING_DATE and TXN_DATE                      *
 *        - New layout with LANE_TXN_ID + POSTED_DATE                      *
 *        - Fee / payment rows (single date + description)                *
 *   3) Keeps date fields as strings (MM/DD/YY)                           *
 *   4) Creates numeric helper columns for amount/balance calculations    *
 *        - amount_num, balance_num                                      *
 *   5) Writes a formatted Excel worksheet                                *
 *        - Freeze header row                                             *
 *        - Auto-size columns                                             *
 *                                                                        *
 * Output                                                                 *
 *   Excel workbook with one sheet: "Transactions"                        *
 *                                                                        *
 * Notes                                                                  *
 *   - Parsing is based on text extraction from PDF pages via pdfplumber. *
 *   - EZ-Pass PDF formats may vary; parsing uses defensive heuristics.   *
 **************************************************************************/
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd
import pdfplumber

# -------------------------------------------------
# Regular expressions used for parsing
# -------------------------------------------------
"""
/**************************************************************************
 * Parsing Regex Patterns                                                 *
 * ---------------------------------------------------------------------- *
 * These patterns help identify tokens in PDF text that represent:         *
 *   - Dates  : MM/DD/YY                                                   *
 *   - Times  : HH:MM                                                      *
 *   - Money  : $1.23 or -$1.23 (with optional commas)                     *
 *   - Tags   : long digit strings representing EZ-Pass tag / plate        *
 **************************************************************************/
"""
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2}$")
TIME_RE = re.compile(r"^\d{2}:\d{2}$")
MONEY_RE = re.compile(r"^-?\$[\d,]+\.\d{2}$")
TAG_RE = re.compile(r"^\d{8,}$")  # EZ-Pass tag / plate number

# -------------------------------------------------
# Main program entry point
# -------------------------------------------------
def main() -> None:
    """
    /**************************************************************************
     * main()                                                                 *
     * ---------------------------------------------------------------------- *
     * High-level flow                                                       *
     *   1) Read command-line arguments (input PDF or folder, output XLSX).   *
     *   2) Collect PDF files to process.                                     *
     *   3) Parse all PDFs -> one combined Transactions DataFrame.            *
     *   4) Export DataFrame to Excel and apply basic formatting.             *
     *                                                                        *
     * Excel formatting                                                      *
     *   - Freeze panes at A2 so the header stays visible.                    *
     *   - Auto-size each column based on up to first 2000 rows.              *
     **************************************************************************/
    """

    # ***********************************************************************
    # Argument Parsing                                                      *
    # ***********************************************************************
    ap = argparse.ArgumentParser()
    ap.add_argument("input", type=str, help="PDF file or folder with PDFs")
    ap.add_argument("output", type=str, help="Output Excel file (.xlsx)")
    args = ap.parse_args()

    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve()

    # ***********************************************************************
    # Input Collection                                                      *
    #   Determine whether input is a single PDF or a directory of PDFs.      *
    # ***********************************************************************
    pdf_files = collect_inputs_func(input_path)

    # ***********************************************************************
    # PDF Parsing                                                           *
    #   Parse all PDFs and combine into one DataFrame.                       *
    # ***********************************************************************
    tx_df = parse_many_func(pdf_files)

    # ***********************************************************************
    # Excel Export                                                          *
    #   Write to "Transactions" sheet, freeze header row, auto-size columns. *
    # ***********************************************************************
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
# Parse one file or many files
# -------------------------------------------------
def parse_many_func(inputs: list[Path]) -> pd.DataFrame:
    """
    /**************************************************************************
     * parse_many(inputs)                                                    *
     * ---------------------------------------------------------------------- *
     * Parses a list of PDF files and concatenates all results into one       *
     * DataFrame.                                                            *
     *                                                                        *
     * Behavior                                                              *
     *   - Each PDF is parsed independently via parse_pdf().                  *
     *   - Results are concatenated with ignore_index=True.                   *
     *   - If no PDFs are provided, returns an empty DataFrame.               *
     **************************************************************************/
    """
    frames = []

    for pdf in inputs:
        frames.append(parse_pdf_func(pdf))

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

# -------------------------------------------------
# Parse a single PDF file
# -------------------------------------------------
def parse_pdf_func(pdf_path: Path) -> pd.DataFrame:
    """
    /**************************************************************************
     * parse_pdf(pdf_path)                                                   *
     * ---------------------------------------------------------------------- *
     * Reads one PDF file and extracts transaction rows from all pages.       *
     *                                                                        *
     * Steps                                                                  *
     *   1) Open PDF with pdfplumber.                                         *
     *   2) For each page: extract text.                                      *
     *   3) Filter page text down to likely transaction lines.                *
     *   4) Parse each line into a normalized row dict.                       *
     *   5) Build a DataFrame from all rows.                                  *
     *   6) Add numeric helper columns for amount and balance.                *
     *                                                                        *
     * Returns                                                                *
     *   pandas DataFrame containing parsed rows for this PDF.                *
     **************************************************************************/
    """
    rows = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in iter_transaction_lines_func(text):
                row = parse_transaction_line_func(line)
                if row:
                    rows.append(row)

    df = pd.DataFrame(rows)

    # ***********************************************************************
    # Keep date columns as strings (MM/DD/YY); ensure missing stays None.   *
    # ***********************************************************************
    for col in ["posting_date", "transaction_date", "entry_date", "exit_date"]:
        if col in df.columns:
            df[col] = df[col].where(df[col].notna(), None)

    # ***********************************************************************
    # Numeric helper columns for calculations / sorting / filtering.        *
    # ***********************************************************************
    if not df.empty:
        df["amount_num"] = df["amount"].apply(money_to_float_func)
        df["balance_num"] = df["balance"].apply(money_to_float_func)

    return df

# -------------------------------------------------
# Extract only the transaction lines from a page
# -------------------------------------------------
def iter_transaction_lines_func(page_text: str) -> Iterable[str]:
    """
    /**************************************************************************
     * iter_transaction_lines(page_text)                                      *
     * ---------------------------------------------------------------------- *
     * Goal                                                                    *
     *   From a page's extracted text, yield only the lines that look like     *
     *   transaction rows (tolls, fees, payments).                             *
     *                                                                          *
     * Supported line styles                                                   *
     *   A) New format (lane txn id first):                                    *
     *        LANE_TXN_ID POSTED_DATE ...                                      *
     *        Example: 31420710413 04/06/25 ...                                *
     *                                                                          *
     *   B) Fee/payment format (date first):                                   *
     *        POSTED_DATE Description AMT BAL                                  *
     *        Example: 04/03/25 Monthly Service Fee -$1.00 -$19.91             *
     *                                                                          *
     *   C) Old format (two dates first):                                      *
     *        POSTING_DATE TXN_DATE ...                                        *
     *        Example: 12/11/24 12/10/24 ...                                   *
     *                                                                          *
     * Filtering strategy                                                      *
     *   - Ignore blank lines.                                                 *
     *   - Require at least a few tokens.                                      *
     *   - Accept lines starting with:                                         *
     *        (digits + date) OR (date)                                        *
     **************************************************************************/
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
def parse_transaction_line_func(line: str) -> Optional[dict]:
    """
    /**************************************************************************
     * parse_transaction_line(line)                                           *
     * ---------------------------------------------------------------------- *
     * Converts one statement row (a single text line) into a structured dict *
     * that matches the unified schema used by the output Excel file.         *
     *                                                                        *
     * Unified Output Schema                                                  *
     *   lane_txn_id, posting_date, transaction_date, tag_or_plate, agency,   *
     *   plaza, entry_date, entry_time, exit_plaza, exit_date, exit_time,     *
     *   plan, vehicle_class, amount, balance, description                    *
     *                                                                        *
     * Supported layouts                                                      *
     *   1) Old toll layout (two dates at start):                             *
     *        POSTING_DATE TXN_DATE TAG AGENCY PLAZA ... PLAN CL AMT BAL       *
     *                                                                        *
     *   2) New toll layout (lane id + date):                                 *
     *        LANE_TXN_ID POSTED_DATE TAG AGENCY PLAZA ENTRY_DATE ENTRY_TIME   *
     *        ... PLAN CL AMT BAL                                              *
     *                                                                        *
     *   3) Fee / payment row (single date at start):                         *
     *        POSTED_DATE Description AMT BAL                                  *
     *                                                                        *
     * Heuristics used                                                        *
     *   - We assume the last two tokens are AMOUNT and BALANCE.              *
     *   - We keep parsing defensive: missing fields remain None.             *
     **************************************************************************/
    """
    tokens = line.split()
    if len(tokens) < 4:
        return None

    # ***********************************************************************
    # Helper: build a full row with all supported columns (default None).   *
    # ***********************************************************************
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

    # ***********************************************************************
    # Basic tail assumption: last two tokens are amount and balance.        *
    # ***********************************************************************
    amount = tokens[-2]
    balance = tokens[-1]

    # ------------------------------------------------------------
    # FORMAT B (new): starts with lane txn id (digits), then a date
    # ------------------------------------------------------------
    if tokens[0].isdigit() and DATE_RE.match(tokens[1]):
        row = empty_row()
        row["lane_txn_id"] = tokens[0]
        row["posting_date"] = tokens[1]

        # If not enough tokens to map to expected fields, keep description
        if len(tokens) < 9:
            # Not enough tokens to be meaningful
            row["amount"] = amount
            row["balance"] = balance
            row["description"] = " ".join(tokens[2:-2]) if len(tokens) > 4 else None
            return row

        # Core fields (best-effort)
        row["tag_or_plate"] = tokens[2]
        row["agency"] = tokens[3] if len(tokens) > 3 else None
        row["plaza"] = tokens[4] if len(tokens) > 4 else None

        # Tail tokens typically: PLAN, CLASS, AMT, BAL
        if len(tokens) >= 6:
            row["plan"] = tokens[-4] if len(tokens) >= 6 else None
            row["vehicle_class"] = tokens[-3] if len(tokens) >= 5 else None

        # Middle = everything between plaza and tail
        middle = tokens[5:-4]

        # Entry date/time is usually first in the middle
        if middle:
            if DATE_RE.match(middle[0]):
                row["entry_date"] = middle[0]
                if len(middle) > 1 and TIME_RE.match(middle[1]):
                    row["entry_time"] = middle[1]
                rest = middle[2:]
            else:
                # If some statements insert extra plaza token, handle gracefully
                rest = middle

            # Optional exit info (rare)
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
    # ------------------------------------------------------------
    if DATE_RE.match(tokens[0]) and len(tokens) > 1 and DATE_RE.match(tokens[1]):
        row = empty_row()
        row["posting_date"] = tokens[0]
        row["transaction_date"] = tokens[1]

        # If 3rd token looks like a tag/plate -> toll row
        # Otherwise it’s fee/payment row
        row["tag_or_plate"] = tokens[2] if len(tokens) > 2 else None

        # Short lines are usually fee/payment-like in older layout
        if len(tokens) <= 6:
            row["description"] = " ".join(tokens[2:-2])
            row["amount"] = amount
            row["balance"] = balance
            return row

        # Best-effort mapping for older layout
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
# Convert money string like "$1.74" to float 1.74
# -------------------------------------------------
def money_to_float_func(value: str) -> Optional[float]:
    """
        /**************************************************************************
         * money_to_float(value)                                                 *
         * ---------------------------------------------------------------------- *
         * Converts a money string from the PDF into a numeric float.             *
         *                                                                        *
         * Examples                                                               *
         *   "$1.74"   ->  1.74                                                   *
         *   "-$6.94"  -> -6.94                                                   *
         *   "$1,234.50" -> 1234.50                                               *
         *                                                                        *
         * Returns                                                                *
         *   float value if conversion succeeds, otherwise None.                  *
         **************************************************************************/
        """
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
# Collect PDF files from input
# -------------------------------------------------
def collect_inputs_func(path: Path) -> list[Path]:
    """
    /**************************************************************************
     * collect_inputs(path)                                                  *
     * ---------------------------------------------------------------------- *
     * Determines which PDF files to parse based on the input path.           *
     *                                                                        *
     * Accepted input                                                        *
     *   - A single PDF file path -> returns [that_file]                      *
     *   - A directory path -> returns all *.pdf inside (sorted)              *
     *                                                                        *
     * Errors                                                                *
     *   - Raises FileNotFoundError if:                                       *
     *        * directory has no PDFs                                         *
     *        * path does not exist                                           *
     **************************************************************************/
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
"""
/**************************************************************************
 * Script Entrypoint                                                      *
 * ---------------------------------------------------------------------- *
 * If this file is executed directly (not imported), run main().           *
 **************************************************************************/
"""
if __name__ == "__main__":
    main()


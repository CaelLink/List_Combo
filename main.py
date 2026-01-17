from pathlib import Path
import pdfplumber
import pandas as pd

# Config
INPUT_DIR = Path("input_pdfs")
OUTPUT_DIR = Path("output")
OUTPUT_FILE = OUTPUT_DIR / "Master_Material_List.xlsx"

NORMALIZE_SPACES = True



# Helpers
HEADER_HINTS = {"quantity", "units", "size", "description"}

# Clean up weird characters, lines, and spaces
def normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = s.strip()
    if NORMALIZE_SPACES:
        s = " ".join(s.split())
    # normalize some common odd spacing characters
    s = s.replace("\u2009", " ").replace("\u00a0", " ")
    s = " ".join(s.split())
    return s

# Create a key for each section, so it knows when categories start and stop
def make_item_key(size: str, description: str, units: str) -> str:
    size_n = normalize_text(size).lower()
    desc_n = normalize_text(description).lower()
    units_n = normalize_text(units).lower()
    return f"{units_n} | {size_n} | {desc_n}"

# Take rows from each table
def extract_rows_from_table(table, source_name):
    rows = []

    if not table or len(table) < 2:
        return rows

    header = [normalize_text(x).lower() for x in table[0]]
    # Check if table looks like column headers
    if not any(h in HEADER_HINTS for h in header):
        return rows

    # Map expected columns by name- handles minor header variations
    def find_col(name):
        for i, h in enumerate(header):
            if name in h:
                return i
        return None

    idx_qty = find_col("quantity")
    idx_units = find_col("units")
    idx_size = find_col("size")
    idx_desc = find_col("description")

    if None in (idx_qty, idx_units, idx_size, idx_desc):
        return rows
    # Iterate through table rows
    for r in table[1:]:
        if not r or len(r) < max(idx_qty, idx_units, idx_size, idx_desc) + 1:
            continue

        qty_raw = normalize_text(r[idx_qty])
        units = normalize_text(r[idx_units])
        size = normalize_text(r[idx_size])
        desc = normalize_text(r[idx_desc])

        # Skip blank lines
        if not qty_raw or not units or not desc:
            continue

        # Convert quantity to float to ensure proper manipulation
        try:
            qty = float(qty_raw)
        except ValueError:
            continue

        rows.append(
            {
                "source": source_name,
                "quantity": qty,
                "units": units,
                "size": size,
                "description": desc,
                "item_key": make_item_key(size, desc, units),
            }
        )
    return rows


# Handle multi-page .pdf file
def stitch_wrapped_lines(lines: list[str]) -> list[str]:

    stitched = []
    i = 0

    while i < len(lines):
        cur = normalize_text(lines[i])
        if not cur:
            i += 1
            continue

        parts = cur.split()

        # Detect a wrapped line
        if len(parts) == 2:
            qty_raw, unit_raw = parts
            try:
                float(qty_raw)
                if unit_raw.upper() in {"EA", "LF"} and i + 1 < len(lines):
                    nxt = normalize_text(lines[i + 1])
                    if nxt:
                        stitched.append(f"{cur} {nxt}")
                        i += 2
                        continue
            except ValueError:
                pass

        stitched.append(cur)
        i += 1

    return stitched

# Parse lines and separate columns from each row of data
def extract_rows_from_text(text, source_name):

    rows = []
    if not text:
        return rows

    # Words/phrases that mark the start of the description portion
    DESC_START_TOKENS = [
        "type", "propress", "wrot", "threaded", "butterfly", "bolts",
        "valve", "adapter", "coupling", "cap", "tee", "ell", "reducer",
        "flange", "plug", "tube", "street", "measurement/balancing"
    ]

    # Return index in tokens where description likely begins
    def find_desc_start_index(tokens):

        for i, t in enumerate(tokens):
            tt = t.lower()
            if any(tt == k or tt.startswith(k) for k in DESC_START_TOKENS):
                return i
        return None

    raw_lines = text.splitlines()
    lines = stitch_wrapped_lines(raw_lines)

    for line in lines:
        line = normalize_text(line)
        if not line:
            continue
        low = line.lower()
        if low.startswith("dkc -") or low.startswith("quantity units"):
            continue

        # Split lines into parts
        parts = line.split()
        if len(parts) < 4:
            continue

        qty_raw, units_raw = parts[0], parts[1]
        try:
            qty = float(qty_raw)
        except ValueError:
            continue

        units = normalize_text(units_raw)
        if units.upper() not in {"EA", "LF"}:
            continue

        remainder = parts[2:]

        desc_start = find_desc_start_index(remainder)

        # If can't find a description start token, skip (or treat all as description)
        if desc_start is None or desc_start == 0:

            size = ""
            desc = " ".join(remainder)
        else:
            size = " ".join(remainder[:desc_start])
            desc = " ".join(remainder[desc_start:])

        size = normalize_text(size)
        desc = normalize_text(desc)

        if not desc:
            continue

        rows.append(
            {
                "source": source_name,
                "quantity": qty,
                "units": units,
                "size": size,
                "description": desc,
                "item_key": make_item_key(size, desc, units),
            }
        )

    return rows

# Returns a list of rows extracted from one PDF.
def extract_pdf(pdf_path: Path):
    source_name = pdf_path.stem
    all_rows = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        total_pages = len(pdf.pages)

        for page_num, page in enumerate(pdf.pages, start=1):
            print(f"  Page {page_num}/{total_pages}")

            # Try table extraction
            table = page.extract_table()
            table_rows = extract_rows_from_table(table, source_name)
            all_rows.extend(table_rows)

            # Fallback to text parsing if table extraction looks empty/weak
            if len(table_rows) < 2:
                page_text = page.extract_text() or ""
                text_rows = extract_rows_from_text(page_text, source_name)
                all_rows.extend(text_rows)

    return all_rows

# Convert size strings to readable numbers
def size_to_float(size: str) -> float:

    if not size:
        return 0.0

    s = size.replace("¼", ".25").replace("½", ".5").replace("¾", ".75")

    # Take first number before 'x' or space
    s = s.split("x")[0].strip()
    s = s.split()[0]

    try:
        return float(s)
    except ValueError:
        return 0.0

# Main
def main():
    # Check input folder exsists
    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"Missing folder: {INPUT_DIR.resolve()}")

    # Create output folder if doesn't exsist
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Open .pdfs
    pdf_files = sorted(INPUT_DIR.glob("*.pdf"))
    if not pdf_files:
        raise FileNotFoundError(f"No PDFs found.")

    # Extract rows from each .pdf
    rows = []
    for pdf_path in pdf_files:
        print(f"Reading: {pdf_path.name}")
        rows.extend(extract_pdf(pdf_path))

    if not rows:
        raise RuntimeError("No rows extracted.")

    # Convert rows into dataframe
    df = pd.DataFrame(rows)

    # Combine everything into one master list
    master = (
        df.groupby(["item_key", "units", "size", "description"], as_index=False)
        .agg(quantity=("quantity", "sum"))
    )

    # LF first, then EA
    unit_order = {"LF": 0, "EA": 1}

    # Temporary sort key so LF appears first
    master["_unit_sort"] = master["units"].str.upper().map(
        lambda u: unit_order.get(u, 99)
    )

    # Temporary numeric size for correct sorting
    master["_size_sort"] = master["size"].apply(size_to_float)

    # Sort: LF first → description A–Z → size ascending
    master = master.sort_values(
        by=["_unit_sort", "description", "_size_sort"],
        ascending=[True, True, True]
    )

    # Remove helper columns
    master = master.drop(columns=["_size_sort","_unit_sort"])

    # Reorder columns
    master = master[["quantity", "units", "size", "description", "item_key"]]


    # Write output Excel file
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        master.to_excel(writer, index=False, sheet_name="Master")
        df.to_excel(writer, index=False, sheet_name="RawExtract")


    print(f"\nDone, Output written to:\n{OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()
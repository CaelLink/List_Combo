# Master Material List Combiner (PDF → Excel)

## What it does
- Reads all PDFs in `input_pdfs/`
- Extracts Quantity / Units / Size / Description
- Combines identical items (Units + Size + Description)
- Sums quantities across PDFs
- Outputs Excel file to `output/Master_Material_List.xlsx`

## One-time setup
1. Install Python 3.10+ from python.org (check "Add Python to PATH")
2. In this folder run:
   pip install -r requirements.txt

## Run for a new project
1. Put PDFs in `input_pdfs/`
2. Run:
   python main.py
   (or double-click run.bat)
3. Output is in:
   output/Master_Material_List.xlsx

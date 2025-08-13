import re
import win32com.client

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

container = session.findById("/app/con[0]/ses[0]/wnd[0]/usr")

LBL_RE = re.compile(r".*/lbl\[(\d+),(\d+)\]$")

def norm_header(s: str) -> str:
    # Normalize headers like "Opret. d." -> "Opret. d." (keep dots), but trim/space-normalize
    return " ".join((s or "").strip().split())

cells = []          # (col:int, row:int, text:str, id:str)
non_table_labels = []

for child in container.Children:
    if "lbl" in child.Id:
        m = LBL_RE.match(child.Id)
        if not m:
            # Label under usr but not in the lbl[i,j] grid -> not part of table
            non_table_labels.append((child.Id, getattr(child, "Text", "").strip()))
            continue
        col, row = int(m.group(1)), int(m.group(2))
        try:
            text = (child.Text or "").strip()
        except Exception:
            text = ""
        cells.append((col, row, text, child.Id))

# 1) Make sure we only have table labels (optional but you asked for "table and nothing else")
if non_table_labels:
    raise RuntimeError(
        "Der findes labels udenfor tabellen (ikke i formatet lbl[col,row]): "
        + ", ".join(f"{i}='{t}'" for i,t in non_table_labels[:10])
        + (" ..." if len(non_table_labels) > 10 else "")
    )

# 2) Build header (row==1) and data rows (row>=3, odd)
headers = {col: norm_header(text) for col, row, text, _ in cells if row == 1}
if not headers:
    raise RuntimeError("Ingen tabel-headers (row=1) fundet.")

data_cells = [(col, row, text) for col, row, text, _ in cells if row >= 3 and row % 2 == 1]

if not data_cells:
    # Still OK if there are zero data rows; we’ll just validate Fejl header exists and is empty-by-definition
    pass

# 3) Sanity checks: rows only 1 or odd >=3 (else it's probably not the table)
unexpected = [(c, r, t) for c, r, t, _ in cells if not (r == 1 or (r >= 3 and r % 2 == 1))]
if unexpected:
    raise RuntimeError(
        "Uventede label-rækker (ikke row=1 eller en ulige række >=3): "
        + ", ".join(f"[{c},{r}]='{t}'" for c, r, t in unexpected[:10])
        + (" ..." if len(unexpected) > 10 else "")
    )

# 4) Order columns and build a column name list
sorted_cols = sorted(headers.keys())
column_names = [headers[c] for c in sorted_cols]

# Find Fejl column (case-insensitive, punctuation tolerant)
def is_fejl(h):
    return h.lower().strip(".:") == "fejl"

fejl_col = None
for c in sorted_cols:
    if is_fejl(headers[c]):
        fejl_col = c
        break

if fejl_col is None:
    raise RuntimeError("Kolonnen 'Fejl' blev ikke fundet i header-rækken.")

# 5) Group data by row index and map to {header: value}
from collections import defaultdict, OrderedDict

rows_by_index = defaultdict(dict)
for col, row, text in data_cells:
    rows_by_index[row][col] = text

# Convert to list of dicts in row order
table_rows = []
for row_idx in sorted(rows_by_index.keys()):
    rowmap = rows_by_index[row_idx]
    record = OrderedDict()
    for c in sorted_cols:
        record[headers[c]] = rowmap.get(c, "")
    table_rows.append(record)

# 6) Validate: every Fejl cell must be empty string
bad_fejl = []
for i, rec in enumerate(table_rows, start=1):
    val = (rec.get(headers[fejl_col]) or "").strip()
    if val != "":
        bad_fejl.append((i, val))

if bad_fejl:
    # Build a helpful error message with first few offending rows
    preview = ", ".join(f"række {i}: '{v}'" for i, v in bad_fejl[:10])
    raise RuntimeError(
        f"Fejl-kolonnen skal være tom i alle rækker, men fandt værdier: {preview}"
        + (" ..." if len(bad_fejl) > 10 else "")
    )

# 7) All good — we have a clean table and Fejl is empty everywhere
print("✅ Tabel verificeret (kun grid-labels, korrekt header/data-rækker).")
print(f"Kolonner: {column_names}")
print(f"Antal rækker: {len(table_rows)}")
# Access example: print each row
for idx, rec in enumerate(table_rows,  start=1):
    print(f"Row {idx}: {dict(rec)}")

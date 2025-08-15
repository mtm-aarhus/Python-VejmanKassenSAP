import win32com.client
import re

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

container = session.findById("/app/con[0]/ses[0]/wnd[0]/usr")

# Collect all label texts in order
texts = []
for child in container.Children:
    if "lbl" in child.Id:  # SAP GUI labels usually have 'lbl' in their Id
        try:
            texts.append((child.Id, (child.Text or "").strip()))
        except Exception:
            continue

# Convenience: just the text values
labels = [t for _, t in texts]
print("All label texts combined:\n" + " | ".join(labels))

# Find the split point
try:
    split_idx = labels.index("Række Fejltekst")
except ValueError:
    raise RuntimeError("Kunne ikke finde 'Række Fejltekst' i labels; kan ikke validere vervolgstekster.")

after = labels[split_idx + 1 :]

# Pattern: "KMD Standardordre <digits> gemt"
pat = re.compile(r"^KMD\s+Standardordre\s+(\d+)\s+gemt$", re.IGNORECASE)

standardordre_ids = []
bad_entries = []

for text in after:
    if text == "":  # empty is allowed, skip it
        continue
    m = pat.match(text)
    if m:
        standardordre_ids.append(m.group(1))
    else:
        bad_entries.append(text)

# If any non-empty entry didn't match, that's an error
if bad_entries:
    raise RuntimeError(
        "Uventet tekst efter 'Række Fejltekst' (skal være 'KMD Standardordre <xyz> gemt' eller tom). "
        f"Fandt i stedet: {bad_entries}"
    )

# At this point, everything non-empty was valid and we've captured all xyz values
print(f"Valideret. Fangede {len(standardordre_ids)} Standardordre-id(s): {standardordre_ids}")

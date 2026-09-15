from pathlib import Path
import pandas as pd

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"
KB_FILE = Path(FOLDER) / "Nanananda_Knowledge_Base_V1.xlsx"
SITE = Path(FOLDER) / "website"
SERMONS_FOLDER = SITE / "sermons"

# Ensure target directories exist
SERMONS_FOLDER.mkdir(parents=True, exist_ok=True)

# Read the Sermons sheet and clean up missing values (replace NaN with empty string)
sermons_df = pd.read_excel(KB_FILE, sheet_name="Sermons").fillna("")

# ======================================
# HOMEPAGE
# ======================================

home = """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Nanananda Dhamma Knowledge Base</title>
</head>
<body>

<h1>Nanananda Dhamma Knowledge Base</h1>

<p>170 Sermons indexed.</p>

<h2>Sermons</h2>

<ul>
"""

for _, row in sermons_df.iterrows():
    sermon = str(row["Sermon ID"]).strip()
    home += f'  <li><a href="sermons/{sermon}.html">{sermon}</a></li>\n'

home += """</ul>

</body>
</html>
"""

with open(SITE / "index.html", "w", encoding="utf-8") as f:
    f.write(home)

# ======================================
# SERMON PAGES
# ======================================

for _, row in sermons_df.iterrows():
    sermon = str(row["Sermon ID"]).strip()
    citation = row["Citation / Opening Passage"]
    source = row["Canonical Source (Likely Source)"]
    subject = row["Dhamma Subject / Keywords"]
    notes = row["Notes"]

    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{sermon}</title>
</head>
<body>

<h1>{sermon}</h1>

<h2>Opening Citation</h2>
<p>{citation}</p>

<h2>Canonical Source</h2>
<p>{source}</p>

<h2>Subjects</h2>
<p>{subject}</p>

<h2>Notes</h2>
<p>{notes}</p>

<hr>
<p><a href="../index.html">← Back to Home</a></p>

</body>
</html>
"""

    with open(SERMONS_FOLDER / f"{sermon}.html", "w", encoding="utf-8") as f:
        f.write(html)

print("\nDONE")
print("Homepage created successfully.")
print(f"Sermon pages generated: {len(sermons_df)}")
print("\nWebsite folder location:")
print(SITE)
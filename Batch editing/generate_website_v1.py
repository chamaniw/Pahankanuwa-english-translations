import pandas as pd
from pathlib import Path

# =====================================
# SETTINGS
# =====================================

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

KB_FILE = Path(FOLDER) / "Nanananda_Knowledge_Base_V1.xlsx"

SITE = Path(FOLDER) / "website"

SERMONS_FOLDER = SITE / "sermons"

SITE.mkdir(exist_ok=True)
SERMONS_FOLDER.mkdir(exist_ok=True)

# =====================================
# LOAD DATA
# =====================================

sermons_df = pd.read_excel(
    KB_FILE,
    sheet_name="Sermons"
)

# =====================================
# HOMEPAGE
# =====================================

home_html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Nanananda Dhamma Knowledge Base</title>
</head>
<body>

<h1>Nanananda Dhamma Knowledge Base</h1>

<p>170 Indexed Sermons</p>

<h2>Sermons</h2>

<ul>
"""

for _, row in sermons_df.iterrows():

    sermon = str(row["Sermon ID"])

    home_html += (
        f'<li><a href="sermons/{sermon}.html f'{sermon}</a></li>\n'
    )

home_html += """
</ul>

</body>
</html>
"""

with open(
    SITE / "index.html",
    "w",
    encoding="utf-8"
) as f:

    f.write(home_html)

# =====================================
# SERMON PAGES
# =====================================

for _, row in sermons_df.iterrows():

    sermon = str(row["Sermon ID"])

    citation = str(
        row["Citation / Opening Passage"]
    )

    source = str(
        row["Canonical Source (Likely Source)"]
    )

    subject = str(
        row["Dhamma Subject / Keywords"]
    )

    notes = str(
        row["Notes"]
    )

    sermon_html = f"""
<!DOCTYPE html>
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

<p>
../index.htmlBack to Home</a>
</p>

</body>
</html>
"""

    with open(
        SERMONS_FOLDER / f"{sermon}.html",
        "w",
        encoding="utf-8"
    ) as f:

        f.write(sermon_html)

# =====================================
# FINISHED
# =====================================

print()
print("DONE")
print()
print("Homepage:")
print(SITE / "index.html")
print()
print("Sermon pages created:")
print(len(sermons_df))
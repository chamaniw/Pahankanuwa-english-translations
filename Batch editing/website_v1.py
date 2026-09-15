import pandas as pd
from pathlib import Path

# --------------------
# SETTINGS
# --------------------

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

KB_FILE = Path(FOLDER) / "Nanananda_Knowledge_Base_V1.xlsx"

SITE = Path(FOLDER) / "website_v1"

SERMONS_FOLDER = SITE / "sermons"

SITE.mkdir(
    exist_ok=True
)

SERMONS_FOLDER.mkdir(
    exist_ok=True
)

# --------------------
# LOAD DATA
# --------------------

sermons_df = pd.read_excel(
    KB_FILE,
    sheet_name="Sermons"
)

print()
print("Workbook loaded")

print(
    "Rows:",
    len(sermons_df)
)

print(
    "First sermon:",
    sermons_df.iloc[0]["Sermon ID"]
)
# --------------------
# TEST FIRST 3 SERMONS
# --------------------

test_df = sermons_df.head(3)

for _, row in test_df.iterrows():

    sermon = str(
        row["Sermon ID"]
    )

    print(
        "SERMON:",
        sermon
    )
# --------------------
# CREATE SIMPLE HTML
# --------------------

html = ""

html += "<html>\n"
html += "<body>\n"

html += "<h1>Nanananda Dhamma Knowledge Base</h1>\n"

for _, row in test_df.iterrows():

    sermon = str(
        row["Sermon ID"]
    )

    html += sermon
    html += "<br>\n"

html += "</body>\n"
html += "</html>\n"

with open(
    SITE / "index.html",
    "w",
    encoding="utf-8"
) as f:

    f.write(html)

print()
print("index.html created")
# --------------------
# CREATE SERMON PAGES
# --------------------

for _, row in test_df.iterrows():

    sermon = str(
        row["Sermon ID"]
    )

    citation = str(
        row["Citation / Opening Passage"]
    )

    html = ""

    html += "<html>\n"
    html += "<body>\n"

    html += "<h1>"
    html += sermon
    html += "</h1>\n"

    html += "<h2>Opening Citation</h2>\n"

    html += "<p>"
    html += citation
    html += "</p>\n"

    html += "</body>\n"
    html += "</html>\n"

    with open(
        SERMONS_FOLDER / (sermon + ".html"),
        "w",
        encoding="utf-8"
    ) as f:

        f.write(html)

print()
print("SERMON PAGES CREATED")
test_link = ""

test_link += "<html>\n"
test_link += "<body>\n"

test_link += '<a href="sermons/sermon_001_ENGLISH.html">CLICK ME</a>\nk += "</body>\n"
test_link += "</html>\n"

with open(
    SITE / "linktest.html",
    "w",
    encoding="utf-8"
) as f:

    f.write(test_link)

print("linktest.html created")
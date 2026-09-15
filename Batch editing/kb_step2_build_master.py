import pandas as pd
from pathlib import Path

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

V3_FILE = Path(FOLDER) / "Nanananda_Dhamma_Index_V3.xlsx"
GEMINI_FILE = Path(FOLDER) / "Suttareference filled Gemini.xlsx"

print("Loading workbooks...")

sermons_df = pd.read_excel(
    GEMINI_FILE,
    sheet_name="Sheet1"
)

citations_df = pd.read_excel(
    V3_FILE,
    sheet_name="Unique_Citation_Index"
)

suttas_df = pd.read_excel(
    V3_FILE,
    sheet_name="Unique_Sutta_Index"
)

stats_df = pd.read_excel(
    V3_FILE,
    sheet_name="Statistics"
)
subject_map = {}

for _, row in sermons_df.iterrows():

    subject_text = str(
        row["Dhamma Subject / Keywords"]
    )

    sermon = row["Sermon ID"]

    for subject in subject_text.split(","):

        subject = subject.strip()

        if not subject:
            continue

        if subject not in subject_map:
            subject_map[subject] = set()

        subject_map[subject].add(sermon)

subject_rows = []

for subject, sermons in subject_map.items():

    subject_rows.append({

        "Subject": subject,

        "Sermons": ", ".join(
            sorted(sermons)
        ),

        "Count": len(sermons)

    })

subjects_df = pd.DataFrame(
    subject_rows
)
output = (
    Path(FOLDER)
    / "Nanananda_Knowledge_Base_V1.xlsx"
)

with pd.ExcelWriter(
    output,
    engine="openpyxl"
) as writer:

    sermons_df.to_excel(
        writer,
        sheet_name="Sermons",
        index=False
    )

    citations_df.to_excel(
        writer,
        sheet_name="Citations",
        index=False
    )

    suttas_df.to_excel(
        writer,
        sheet_name="Suttas",
        index=False
    )

    subjects_df.to_excel(
        writer,
        sheet_name="Subjects",
        index=False
    )

    stats_df.to_excel(
        writer,
        sheet_name="Statistics",
        index=False
    )

print()
print("DONE")
print(output)

print()
print("Sermons:", len(sermons_df))
print("Subjects:", len(subjects_df))
print("Citations:", len(citations_df))
print("Suttas:", len(suttas_df))
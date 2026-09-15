from docx import Document
from pathlib import Path
import pandas as pd
import re

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

PALI_MARKS = "āīūṅñṭḍṇḷṃĀĪŪṄÑṬḌṆḶṂ"

rows = []

for file in Path(FOLDER).rglob("*.docx"):

    try:

        sermon = file.stem

        doc = Document(file)

        paragraphs = [
            p.text.strip()
            for p in doc.paragraphs
            if p.text.strip()
        ]

        for para_no, para in enumerate(paragraphs, start=1):

            matches = re.findall(r"\*([^*]+)\*", para)

            for citation in matches:

                citation = citation.strip()

                word_count = len(citation.split())

                if word_count < 3:
                    continue

                if not any(ch in citation for ch in PALI_MARKS):
                    continue

                rows.append({
                    "Sermon": sermon,
                    "Paragraph": para_no,
                    "Citation": citation,
                    "Word Count": word_count,
                    "Type": "Unknown"
                })

    except Exception as e:

        print("ERROR:", file)
        print(e)

df = pd.DataFrame(rows)

df = df.drop_duplicates(
    subset=["Sermon", "Citation"]
)

output = Path(FOLDER) / "V3_Step1.xlsx"

with pd.ExcelWriter(output, engine="openpyxl") as writer:

    df.to_excel(
        writer,
        sheet_name="Citations",
        index=False
    )

print()
print("DONE")
print(output)
print("Rows:", len(df))
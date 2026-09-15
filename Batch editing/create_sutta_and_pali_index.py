from docx import Document
from pathlib import Path
import pandas as pd
import re

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

records = []

for file in Path(FOLDER).rglob("*.docx"):

    try:

        doc = Document(file)

        paragraphs = [
            p.text.strip()
            for p in doc.paragraphs
            if p.text.strip()
        ]

        sermon = file.stem

        for para_no, para in enumerate(paragraphs, start=1):

            matches = re.findall(r"\*([^*]+)\*", para)

            for citation in matches:

                citation = citation.strip()

                # Ignore single-word glossary items
                word_count = len(citation.split())

                if word_count < 3:
                    continue

                records.append({
                    "Sermon": sermon,
                    "Paragraph": para_no,
                    "Pali Citation": citation,
                    "Word Count": word_count,
                    "Likely Source": "",
                    "Subject": "",
                    "Notes": ""
                })

    except Exception as e:

        print(file)
        print(e)

df = pd.DataFrame(records)

df = df.drop_duplicates(
    subset=["Sermon", "Pali Citation"]
)

df = df.sort_values(
    by=["Sermon", "Paragraph"]
)

output = Path(FOLDER) / "Sutta_and_Pali_Index_Workbook.xlsx"

with pd.ExcelWriter(output) as writer:

    df.to_excel(
        writer,
        sheet_name="Pali Citations",
        index=False
    )

print()
print("DONE")
print(output)
print(f"Citations found: {len(df)}")
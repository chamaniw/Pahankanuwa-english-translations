from docx import Document
from pathlib import Path
import pandas as pd
import re

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

PALI_MARKS = "āīūṅñṭḍṇḷṃĀĪŪṄÑṬḌṆḶṂ"

ENGLISH_WORDS = [
    "the",
    "this",
    "that",
    "these",
    "those",
    "because",
    "however",
    "therefore",
    "meaning",
    "means",
    "example",
    "wise attention"
]

good_rows = []
review_rows = []

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
		citation = citation.replace('"', '')
		citation = citation.replace('“', '')
		citation = citation.replace('”', '')
		citation = citation.replace("'", '')
		citation = citation.strip()

                word_count = len(citation.split())

                if word_count < 3:
                    continue

                if not any(ch in citation for ch in PALI_MARKS):
                    continue

                lower = citation.lower()

                reason = ""

                if word_count > 15:
                    reason = "Over 15 words"

                elif "[" in citation or "]" in citation:
    reason = "Contains brackets"

elif any(word in lower for word in ENGLISH_WORDS):
    reason = "Contains English"

else:
    reason = ""
                    reason = "Contains brackets"

                elif "(" in citation or ")" in citation:
                    reason = "Contains parentheses"

                elif any(
                    word in lower
                    for word in ENGLISH_WORDS
                ):
                    reason = "Contains English"

                if reason:

                    review_rows.append({
                        "Sermon": sermon,
                        "Paragraph": para_no,
                        "Content": citation,
                        "Reason": reason
                    })

                else:

                    good_rows.append({
                        "Sermon": sermon,
                        "Paragraph": para_no,
                        "Citation": citation,
                        "Word Count": word_count,
                        "Type": "Unknown",
                        "Likely Source": "",
                        "Subject": "",
                        "Notes": ""
                    })

    except Exception as e:

        print(file)
        print(e)

good_df = pd.DataFrame(good_rows)

good_df = good_df.drop_duplicates(
    subset=["Sermon", "Citation"]
)

review_df = pd.DataFrame(review_rows)

output = Path(FOLDER) / "V3_Step2.xlsx"

with pd.ExcelWriter(
    output,
    engine="openpyxl"
) as writer:

    good_df.to_excel(
        writer,
        sheet_name="Pali_Citations_Clean",
        index=False
    )

    review_df.to_excel(
        writer,
        sheet_name="Review_Required",
        index=False
    )

print()
print("DONE")
print(output)
print("Clean Citations:", len(good_df))
print("Review Items:", len(review_df))
from docx import Document
from pathlib import Path
from collections import defaultdict
import pandas as pd
import re

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

OUTPUT_FILE = "Nanananda_Dhamma_Index_V3.xlsx"

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

SUTTA_PATTERNS = [
    "Dhammapada",
    "Udāna",
    "Itivuttaka",
    "Sutta Nipāta",
    "Saṁyutta Nikāya",
    "Majjhima Nikāya",
    "Dīgha Nikāya",
    "Aṅguttara Nikāya",
    "Dhammacakkappavattana Sutta",
    "Anattalakkhaṇa Sutta",
    "Satipaṭṭhāna Sutta",
    "Mahāparinibbāna Sutta",
    "Madhupiṇḍika Sutta",
    "Kaccānagotta Sutta",
    "Mahānidāna Sutta"
]

citation_rows = []
review_rows = []
reference_rows = []
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
                citation = citation.replace("'", "")

                word_count = len(citation.split())

                if word_count < 3:
                    continue

                if not any(ch in citation for ch in PALI_MARKS):
                    continue

                lower = citation.lower()

                if (
                    "[" in citation
                    or "]" in citation
                    or any(word in lower for word in ENGLISH_WORDS)
                ):

                    review_rows.append({
                        "Sermon": sermon,
                        "Paragraph": para_no,
                        "Content": citation,
                        "Reason": "Review Required"
                    })

                    continue

                if citation.count(",") >= 2:
                    citation_type = "Doctrine"

                elif word_count > 15:
                    citation_type = "Long Quotation"

                else:
                    citation_type = "Verse"

                citation_rows.append({
                    "Sermon": sermon,
                    "Paragraph": para_no,
                    "Citation": citation,
                    "Word Count": word_count,
                    "Type": citation_type,
                    "Likely Source": "",
                    "Subject": "",
                    "Notes": ""
                })

            for pattern in SUTTA_PATTERNS:

                if pattern.lower() in para.lower():

                    reference_rows.append({
                        "Sermon": sermon,
                        "Paragraph": para_no,
                        "Reference": pattern,
                        "Reference Type": "Unknown",
                        "Notes": ""
                    })

    except Exception as e:

        print("ERROR:", file)
        print(e)
cit_df = pd.DataFrame(citation_rows)

if not cit_df.empty:

    cit_df = cit_df.drop_duplicates(
        subset=["Sermon", "Citation"]
    )

ref_df = pd.DataFrame(reference_rows)

if not ref_df.empty:

    ref_df = ref_df.drop_duplicates()

review_df = pd.DataFrame(review_rows)

citation_freq_df = (
    cit_df["Citation"]
    .value_counts()
    .reset_index()
)

citation_freq_df.columns = [
    "Citation",
    "Count"
]

reference_freq_df = (
    ref_df["Reference"]
    .value_counts()
    .reset_index()
)

reference_freq_df.columns = [
    "Reference",
    "Count"
]
citation_index = defaultdict(set)

for _, row in cit_df.iterrows():

    citation_index[row["Citation"]].add(
        row["Sermon"]
    )

unique_citation_rows = []

for citation, sermons in citation_index.items():

    unique_citation_rows.append({
        "Citation": citation,
        "Sermons": ", ".join(sorted(sermons))
    })

unique_citation_df = pd.DataFrame(
    unique_citation_rows
)

sutta_index = defaultdict(set)

for _, row in ref_df.iterrows():

    sutta_index[row["Reference"]].add(
        row["Sermon"]
    )

unique_sutta_rows = []

for ref, sermons in sutta_index.items():

    unique_sutta_rows.append({
        "Reference": ref,
        "Sermons": ", ".join(sorted(sermons))
    })

unique_sutta_df = pd.DataFrame(
    unique_sutta_rows
)
stats_df = pd.DataFrame({

    "Metric": [
        "Total Sermons",
        "Total Citations",
        "Unique Citations",
        "Verses",
        "Doctrines",
        "Long Quotations",
        "Review Items",
        "Total References",
        "Unique References"
    ],

    "Value": [
        cit_df["Sermon"].nunique(),
        len(cit_df),
        cit_df["Citation"].nunique(),
        len(cit_df[cit_df["Type"]=="Verse"]),
        len(cit_df[cit_df["Type"]=="Doctrine"]),
        len(cit_df[cit_df["Type"]=="Long Quotation"]),
        len(review_df),
        len(ref_df),
        ref_df["Reference"].nunique()
    ]
})

output = Path(FOLDER) / OUTPUT_FILE

with pd.ExcelWriter(output, engine="openpyxl") as writer:

    cit_df.to_excel(writer, sheet_name="Pali_Citations", index=False)

    ref_df.to_excel(writer, sheet_name="Sutta_References", index=False)

    citation_freq_df.to_excel(writer, sheet_name="Citation_Frequency", index=False)

    reference_freq_df.to_excel(writer, sheet_name="Reference_Frequency", index=False)

    unique_citation_df.to_excel(writer, sheet_name="Unique_Citation_Index", index=False)

    unique_sutta_df.to_excel(writer, sheet_name="Unique_Sutta_Index", index=False)

    review_df.to_excel(writer, sheet_name="Review_Required", index=False)

    stats_df.to_excel(writer, sheet_name="Statistics", index=False)

print()
print("DONE")
print(output)
print()
print("Total Citations:", len(cit_df))
print("Unique Citations:", cit_df["Citation"].nunique())
print("Review Items:", len(review_df))
print("Total References:", len(ref_df))
print("Unique References:", ref_df["Reference"].nunique())

from docx import Document
from pathlib import Path
from collections import defaultdict
import pandas as pd
import re

# ==================================================
# CONFIGURATION
# ==================================================

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

OUTPUT_FILE = "Sutta_and_Pali_Index_Workbook_V3.xlsx"

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
    "wise attention",
    "for instance",
]

SUTTA_PATTERNS = [
    r"Dhammacakkappavattana Sutta",
    r"Anattalakkhaṇa Sutta",
    r"Satipaṭṭhāna Sutta",
    r"Mahāparinibbāna Sutta",
    r"Madhupiṇḍika Sutta",
    r"Mahāsaḷāyatanika Sutta",
    r"Aniccatā Sutta",
    r"Cūḷagosinga Sutta",
    r"Upakkilesa Sutta",
    r"Mahānidāna Sutta",
    r"Kaccānagotta Sutta",
    r"Dhammapada",
    r"Udāna",
    r"Itivuttaka",
    r"Saṁyutta Nikāya",
    r"Majjhima Nikāya",
    r"Dīgha Nikāya",
    r"Aṅguttara Nikāya",
    r"Sutta Nipāta",
]
# ==================================================
# STORAGE
# ==================================================

citations = []
references = []
review_rows = []

# ==================================================
# HELPERS
# ==================================================

def contains_pali_marks(text):
    return any(ch in text for ch in PALI_MARKS)


def contains_english(text):
    lower = text.lower()

    return any(word in lower for word in ENGLISH_WORDS)


def looks_suspicious(text, wc):

    if wc > 15:
        return True

    if "[" in text or "]" in text:
        return True

    if "(" in text or ")" in text:
        return True

    if contains_english(text):
        return True

    return False


# ==================================================
# MAIN EXTRACTION
# ==================================================

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

            # ------------------------------------------
            # PALI CITATIONS
            # ------------------------------------------

            matches = re.findall(r"\*([^*]+)\*", para)

            for citation in matches:

                citation = citation.strip()

                wc = len(citation.split())

                # Too short
                if wc < 3:
                    continue

                # Suspicious -> Review sheet
                if looks_suspicious(citation, wc):

                    review_rows.append({
                        "Sermon": sermon,
                        "Paragraph": para_no,
                        "Content": citation,
                        "Reason": "Needs review"
                    })

                    continue

                # Must contain Pali diacritics
                if not contains_pali_marks(citation):
                    continue

                citations.append({
                    "Sermon": sermon,
                    "Paragraph": para_no,
                    "Pali Citation": citation,
                    "Word Count": wc,
                    "Type": "Unknown",
                    "Likely Source": "",
                    "Subject": "",
                    "Notes": ""
                })

            # ------------------------------------------
            # SUTTA REFERENCES
            # ------------------------------------------

            for pattern in SUTTA_PATTERNS:

                found = re.findall(
                    pattern,
                    para,
                    flags=re.IGNORECASE
                )

                for item in found:

                    references.append({
                        "Sermon": sermon,
                        "Paragraph": para_no,
                        "Reference": item,
                        "Reference Type": "Unknown",
                        "Notes": ""
                    })

    except Exception as e:

        print(f"ERROR: {file}")
        print(e)

# ==================================================
# DATAFRAMES
# ==================================================

cit_df = pd.DataFrame(citations)

if not cit_df.empty:

    cit_df = cit_df.drop_duplicates(
        subset=["Sermon", "Pali Citation"]
    )

ref_df = pd.DataFrame(references)

if not ref_df.empty:

    ref_df = ref_df.drop_duplicates()

review_df = pd.DataFrame(review_rows)

# ==================================================
# CITATION FREQUENCY
# ==================================================

if not cit_df.empty:

    citation_freq = cit_df["Pali Citation"].value_counts().reset_index()

    citation_freq.columns = [
        "Citation",
        "Count"
    ]

else:

    citation_freq = pd.DataFrame()

# ==================================================
# REFERENCE FREQUENCY
# ==================================================

if not ref_df.empty:

    reference_freq = ref_df["Reference"].value_counts().reset_index()

    reference_freq.columns = [
        "Reference",
        "Count"
    ]

else:

    reference_freq = pd.DataFrame()

# ==================================================
# UNIQUE CITATION INDEX
# ==================================================

citation_index = defaultdict(set)

if not cit_df.empty:

    for _, row in cit_df.iterrows():

        citation_index[row["Pali Citation"]].add(
            row["Sermon"]
        )

citation_index_rows = []

for citation, sermons in citation_index.items():

    citation_index_rows.append({
        "Citation": citation,
        "Sermons": ", ".join(sorted(sermons))
    })

citation_index_df = pd.DataFrame(citation_index_rows)

# ==================================================
# UNIQUE SUTTA INDEX
# ==================================================

sutta_index = defaultdict(set)

if not ref_df.empty:

    for _, row in ref_df.iterrows():

    
from docx import Document
from pathlib import Path
from collections import defaultdict
import pandas as pd
import re

# ==========================================
# SETTINGS
# ==========================================

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
    "wise attention"
]

SUTTA_PATTERNS = [
    "Dhammacakkappavattana Sutta",
    "Anattalakkhaṇa Sutta",
    "Satipaṭṭhāna Sutta",
    "Mahāparinibbāna Sutta",
    "Madhupiṇḍika Sutta",
    "Mahāsaḷāyatanika Sutta",
    "Aniccatā Sutta",
    "Cūḷagosinga Sutta",
    "Upakkilesa Sutta",
    "Mahānidāna Sutta",
    "Kaccānagotta Sutta",
    "Dhammapada",
    "Udāna",
    "Itivuttaka",
    "Saṁyutta Nikāya",
    "Majjhima Nikāya",
    "Dīgha Nikāya",
    "Aṅguttara Nikāya",
    "Sutta Nipāta"
]

# ==========================================
# STORAGE
# ==========================================

citation_rows = []
reference_rows = []
review_rows = []

# ==========================================
# HELPER FUNCTIONS
# ==========================================

def contains_pali(text):
    return any(ch in text for ch in PALI_MARKS)

def contains_english(text):
    txt = text.lower()
    return any(word in txt for word in ENGLISH_WORDS)

# ==========================================
# PROCESS DOCX FILES
# ==========================================

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

            # ==========================
            # PALI CITATIONS
            # ==========================

            matches = re.findall(r"\*([^*]+)\*", para)

            for citation in matches:

                citation = citation.strip()

                word_count = len(citation.split())

                if word_count < 3:
                    continue

                if not contains_pali(citation):
                    continue

                suspicious = (
                    word_count > 15
                    or contains_english(citation)
                    or "[" in citation
                    or "]" in citation
                    or "(" in citation
                    or ")" in citation
                )
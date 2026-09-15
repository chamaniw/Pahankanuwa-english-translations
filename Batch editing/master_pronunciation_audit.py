from docx import Document
from pathlib import Path
from collections import Counter
import pandas as pd
import re

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

counter = Counter()

# Diacritics commonly found in Pāli
PALI_MARKS = "āīūṅñṭḍṇḷṃĀĪŪṄÑṬḌṆḶṂ"

for file in Path(FOLDER).rglob("*.docx"):

    try:

        doc = Document(file)

        text = "\n".join(p.text for p in doc.paragraphs)

        # 1. Words containing diacritics
        words = re.findall(r"\b[\w\-]+\b", text)

        for word in words:

            if any(ch in word for ch in PALI_MARKS):

                counter[word] += 1

        # 2. Italicised/glossary style words
        starred = re.findall(r"\*([^*]+)\*", text)

        for word in starred:

            word = word.strip()

            if len(word) > 1:

                counter[word] += 1

        # 3. Hyphenated doctrinal terms
        compounds = re.findall(
            r"\b[A-Za-zĀĪŪṄÑṬḌṆḶṂāīūṅñṭḍṇḷṃ]+-[A-Za-zĀĪŪṄÑṬḌṆḶṂāīūṅñṭḍṇḷṃ\-]+\b",
            text
        )

        for word in compounds:

            counter[word] += 1

    except Exception as e:

        print(file)
        print(e)

rows = []

for term, count in counter.most_common():

    rows.append({
        "Term": term,
        "Count": count,
        "Pronunciation": "",
        "Notes": ""
    })

df = pd.DataFrame(rows)

output = Path(FOLDER) / "Master_Pronunciation_Audit.xlsx"

df.to_excel(output, index=False)

print("DONE")
print(output)
print(f"Terms found: {len(df)}")
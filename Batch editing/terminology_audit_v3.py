from docx import Document
from pathlib import Path
import pandas as pd
from collections import Counter
import re

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

TERMS = [
    "Nibbāna",
    "Anattā",
    "Anicca",
    "Dukkha",
    "Taṇhā",
    "Vedanā",
    "Saṅkhāra",
    "Bhava",
    "Jāti",
    "Yoniso manasikāra"
]

summary_counter = Counter()
details = []

for file in Path(FOLDER).rglob("*.docx"):

    try:
        doc = Document(file)

        text = "\n".join(p.text for p in doc.paragraphs)

        for term in TERMS:

            patterns = [

                rf"\*?{re.escape(term)}\*?\s*\([^)]+\)",

                rf"\*?{re.escape(term)}\*?\s*\[[^\]]+\]",

                rf"\*?{re.escape(term)}\*?\s*[:,-]\s*[^\n\r]+"

            ]

            for pattern in patterns:

                matches = re.findall(
                    pattern,
                    text,
                    flags=re.IGNORECASE
                )

                for match in matches:

                    phrase = match.strip()

                    summary_counter[phrase] += 1

                    details.append({
                        "Term": term,
                        "Phrase": phrase,
                        "File": file.name
                    })

    except Exception as e:

        print(file)
        print(e)

summary = pd.DataFrame(
    [
        {
            "Phrase": phrase,
            "Count": count
        }
        for phrase, count
        in summary_counter.most_common()
    ]
)

detail_df = pd.DataFrame(details)

output = Path(FOLDER) / "Terminology_Audit_v3.xlsx"

with pd.ExcelWriter(output) as writer:

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    detail_df.to_excel(
        writer,
        sheet_name="Occurrences",
        index=False
    )

print()
print("DONE")
print(output)
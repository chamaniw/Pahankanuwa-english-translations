from docx import Document
from pathlib import Path
import pandas as pd
import re

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

SEARCH_WORD = "buddha"

results = []

print("Scanning documents...")

for file in Path(FOLDER).rglob("*.docx"):

    try:
        doc = Document(file)

        text = " ".join(p.text for p in doc.paragraphs)

        words = re.findall(r"\b[\w'-]+\b", text)

        for i, word in enumerate(words):

            if word.lower() == SEARCH_WORD.lower():

                before = words[max(0, i-10):i]
                after = words[i+1:i+11]

                results.append({
                    "File": str(file.relative_to(FOLDER)),
                    "Context":
                    f"{' '.join(before)} >>> {word} <<< {' '.join(after)}"
                })

    except Exception as e:
        print(f"Could not read {file}: {e}")

df = pd.DataFrame(results)

output_file = Path(FOLDER) / "Buddha_Context_Report.xlsx"

df.to_excel(output_file, index=False)

print()
print("Done!")
print(f"Found {len(results)} occurrences.")
print(f"Report saved to:")
print(output_file)
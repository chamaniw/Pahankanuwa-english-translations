from docx import Document
from pathlib import Path
import pandas as pd
import re

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

results = []

for file in Path(FOLDER).rglob("*.docx"):

    try:

        doc = Document(file)

        text = " ".join(p.text for p in doc.paragraphs)

        words = re.findall(r"[\w'-]+", text)

        for i, word in enumerate(words):

            if word.lower() in ["buddha", "father"]:

                start = max(0, i - 6)
                end = min(len(words), i + 7)

                window = words[start:end]

                lower_window = [w.lower() for w in window]

                if "buddha" in lower_window and "father" in lower_window:

                    results.append({
                        "File": file.name,
                        "Context":
                        " ".join(window)
                    })

    except Exception as e:

        print(file)
        print(e)

df = pd.DataFrame(results)

output = Path(FOLDER) / "Buddha_Father_Proximity_Audit.xlsx"

df.to_excel(output, index=False)

print()
print("DONE")
print(output)
print(f"Occurrences found: {len(df)}")
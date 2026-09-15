from docx import Document
from pathlib import Path
import pandas as pd

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

records = []

for file in Path(FOLDER).rglob("*.docx"):

    try:

        doc = Document(file)

        paras = []

        for p in doc.paragraphs:
            text = p.text.strip()

            if text:
                paras.append(text)

        row = {
            "File": file.name
        }

        # Original paragraphs
        for i in range(20):

            value = paras[i] if i < len(paras) else ""

            row[f"Original {i+1}"] = value

        # New paragraphs initially identical
        for i in range(20):

            value = paras[i] if i < len(paras) else ""

            row[f"New {i+1}"] = value

        records.append(row)

    except Exception as e:
        print(f"Error reading {file}")
        print(e)

df = pd.DataFrame(records)

output = Path(FOLDER) / "Sermon_Editing_Workbook.xlsx"

df.to_excel(output, index=False)

print("Done")
print(output)
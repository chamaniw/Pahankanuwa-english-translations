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

        for i in range(20):
            row[f"Paragraph {i+1}"] = paras[i] if i < len(paras) else ""

        records.append(row)

    except Exception as e:
        print(f"Error reading {file}")
        print(e)

df = pd.DataFrame(records)

output = Path(FOLDER) / "Opening_Paragraphs.xlsx"

df.to_excel(output, index=False)

print("Done")
print(f"Saved to: {output}")
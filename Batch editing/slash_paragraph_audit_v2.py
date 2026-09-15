from docx import Document
from pathlib import Path
import pandas as pd

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

results = []

for file in Path(FOLDER).rglob("*.docx"):

    try:

        doc = Document(file)

        for i, para in enumerate(doc.paragraphs, start=1):

            text = para.text.strip()

            if "/" in text:

                results.append({
                    "File": file.name,
                    "Paragraph No": i,
                    "Paragraph Text": text
                })

    except Exception as e:

        print(f"Error reading {file}")
        print(e)

df = pd.DataFrame(results)

df = df.sort_values(
    by=["File", "Paragraph No"]
)

output = Path(FOLDER) / "Slash_Paragraph_Audit_v2.xlsx"

df.to_excel(output, index=False)

print()
print("DONE")
print(f"Paragraphs found: {len(df)}")
print(f"Saved to:")
print(output)
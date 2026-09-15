from docx import Document
from pathlib import Path
import pandas as pd

FOLDER = Path(r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations")

WORKBOOK = FOLDER / "Front_Page_Editing_Workbook.xlsx"

df = pd.read_excel(WORKBOOK)

changes = []

for _, row in df.iterrows():

    filename = row["File"]

    file_path = FOLDER / filename

    if not file_path.exists():
        print(f"Missing file: {filename}")
        continue

    try:

        doc = Document(file_path)

        modified = False

        for i in range(1, 11):

            original = row.get(f"Original {i}", "")
            new = row.get(f"New {i}", "")

            if pd.isna(original):
                original = ""

            if pd.isna(new):
                new = ""

            original = str(original).strip()
            new = str(new).strip()

            if original == new:
                continue

            for para in doc.paragraphs:

                if para.text.strip() == original:

                    para.text = new

                    modified = True

                    changes.append({
                        "File": filename,
                        "Paragraph": i,
                        "Old": original,
                        "New": new
                    })

                    break

        if modified:

            doc.save(file_path)

            print(f"Updated: {filename}")

    except Exception as e:

        print(f"Error processing {filename}")
        print(e)

report = pd.DataFrame(changes)

report_file = FOLDER / "Front_Page_Edits_Report.xlsx"

report.to_excel(report_file, index=False)

print()
print("DONE")
print(f"Changes applied: {len(changes)}")
print(f"Report: {report_file}")
import win32com.client as win32
from pathlib import Path
import pandas as pd

FOLDER = Path(r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations")

word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False

report = []

for file in FOLDER.rglob("*.docx"):

    try:

        doc = word.Documents.Open(str(file))

        result = doc.Content.Find.Execute(
            FindText="*Vedanā* [Feeling]",
            ReplaceWith="*Vedanā* (Feeling)",
            Replace=2
        )

        if result:

            report.append({
                "File": file.name
            })

            print(f"Updated: {file.name}")

        doc.Save()
        doc.Close()

    except Exception as e:

        print(file.name)
        print(e)

word.Quit()

pd.DataFrame(report).to_excel(
    FOLDER / "Vedana_Fix_Report.xlsx",
    index=False
)

print("DONE")
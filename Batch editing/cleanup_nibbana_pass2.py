import win32com.client as win32
from pathlib import Path
import pandas as pd

FOLDER = Path(r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations")

replacements = {

"Nibbāna* (Extinction / the Mind Stilled)": "Nibbāna",
"Nibbāna* (Extinguishment)": "Nibbāna",
"Nibbāna* (the Mind Stilled / extinguishment)": "Nibbāna",
"Nibbāna (Extinction / The Mind Stilled)": "Nibbāna",
"Nibbāna (The Mind Stilled)": "Nibbāna",
"Nibbāna (the mind stilled)": "Nibbāna",
"Nibbāna* (Extinction/Extinguishment)": "Nibbāna",
"Nibbāna* (the mind stilled / extinction)": "Nibbāna",
"Nibbāna* (The Mind Stilled / Extinguishment)": "Nibbāna",
"Nibbāna* (the Mind Stilled / Extinguishment)": "Nibbāna",
"Nibbāna* (Extinction)": "Nibbāna",
"Nibbāna* (the Mind Stilled)": "Nibbāna",
"Nibbāna* - The Mind Stilled)": "Nibbāna",
"Nibbāna* (The Mind Stilled / Extinction)": "Nibbāna",
"Nibbāna (Extinguishment)": "Nibbāna",
"Nibbāna* (the Mind Stilled / Extinction / Extinguishment)": "Nibbāna",
"Nibbāna** (**The Mind Stilled / Extinction / Extinguishment**)": "Nibbāna",
"Nibbāna* (the Mind Stilled / Extinction)": "Nibbāna",
"Nibbāna (the mind stilled / extinguishment)": "Nibbāna",
"Nibbāna* (the mind stilled / extinction / extinguishment)": "Nibbāna",
"Nibbāna* (the mind stilled / extinguishment)": "Nibbāna",
"Nibbāna (Extinction / Extinguishment)": "Nibbāna",
"Nibbāna* (The Mind Stilled / Extinction / Extinguishment)": "Nibbāna",
"Nibbāna*)": "Nibbāna",
"Nibbāna (the Mind Stilled)": "Nibbāna",
"Nibbāna* (The Mind Stilled)": "Nibbāna"

}

word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False

report = []

for file in FOLDER.rglob("*.docx"):

    try:

        doc = word.Documents.Open(str(file))

        changes = 0

        for find_text, replace_text in replacements.items():

            result = doc.Content.Find.Execute(
                FindText=find_text,
                ReplaceWith=replace_text,
                Replace=2
            )

            if result:
                changes += 1

        if changes:

            report.append({
                "File": file.name,
                "Replacement patterns found": changes
            })

            print(f"Updated: {file.name}")

        doc.Save()
        doc.Close()

    except Exception as e:
        print(file.name)
        print(e)

word.Quit()

pd.DataFrame(report).to_excel(
    FOLDER / "Nibbana_Cleanup_Pass2_Report.xlsx",
    index=False
)

print()
print("DONE")
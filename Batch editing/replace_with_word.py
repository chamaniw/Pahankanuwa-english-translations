import win32com.client as win32
from pathlib import Path
import pandas as pd

FOLDER = Path(r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations")

records = []

word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False

for file in FOLDER.rglob("*.docx"):

    try:

        print(f"Processing: {file.name}")

        doc = word.Documents.Open(str(file))

        # Count before replacement
        rng = doc.Content

        count1 = 0
        find = rng.Find

        find.Text = "Father Buddha"

        while find.Execute():
            count1 += 1
            rng.Start = rng.End

        rng = doc.Content

        count2 = 0
        find = rng.Find

        find.Text = "Buddha Father"

        while find.Execute():
            count2 += 1
            rng.Start = rng.End

        # Replace Father Buddha
        doc.Content.Find.Execute(
            FindText="Father Buddha",
            ReplaceWith="Buddha",
            Replace=2
        )

        # Replace Buddha Father
        doc.Content.Find.Execute(
            FindText="Buddha Father",
            ReplaceWith="Buddha",
            Replace=2
        )

        total = count1 + count2

        if total > 0:
            records.append({
                "File": file.name,
                "Father Buddha": count1,
                "Buddha Father": count2,
                "Total": total
            })

        doc.Save()
        doc.Close()

    except Exception as e:
        print(file)
        print(e)

word.Quit()

df = pd.DataFrame(records)

report = FOLDER / "Replacement_Report.xlsx"

df.to_excel(report, index=False)

print()
print("DONE")
print(report)
from docx import Document
from pathlib import Path
import pandas as pd

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

results = []

for file in Path(FOLDER).rglob("*.docx"):

    try:

        doc = *ocument(file)

        text = "\n"*join(p.text for p in doc.paragraph*)

        slash_count = text.coun*("/")

        results.append({
  *         "File": file.name,
      *     "Slash Count": slash_count
  *     })

    except*Exception as e:

        print*f"Error reading {file}")
        p*int(e)

df = pd.DataFrame(results)*
df = df.sort_values(
    by="Slas* Count",
    ascending=False
)

ou*put = Path(FOLDER) / "Slash_Audit.*lsx"

df.to_excel(output, index=Fa*se)

print()
print("DONE")
print(o*tput)
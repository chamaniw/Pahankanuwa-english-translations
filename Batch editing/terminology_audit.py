from docx import Document
from pathlib import Path
import pandas as pd
from collections import defaultdict

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

TERMS = [

    "*Saṅkhāra*",
    "*Vedanā*",
    "*Taṇhā*",
    "*Upādāna*",
    "*Bhava*",
    "*Jāti*",
    "*Jarāmaraṇa*",
    "*Anattā*",
    "*Anicca*",
    "*Dukkha*",
    "*Yoniso manasikāra*",
    "*Sakkāya-diṭṭhi*",
    "*Anusaya*",
    "*Asaṅkhata*",
    "*Māyā*",
    "*Vipassanā*",
    "*Avijjā*",
    "*Paṭicca-samuppāda*",
    "*Kalyāṇamitta*",
    "Nibbāna"
]

records = []

for file in Path(FOLDER).rglob("*.docx"):

    try:

        doc = Document(file)

        text = "\n".join(p.text for p in doc.paragraphs)

        for term in TERMS:

            pos = 0

            while True:

                pos = text.find(term, pos)

                if pos == -1:
                    break

                end = text.find(")", pos)

                if end == -1:
                    end = pos + 150

                phrase = text[pos:end+1].strip()

                records.append({
                    "Term": term,
                    "Phrase Found": phrase,
                    "File": file.name
                })

                pos += len(term)

    except Exception as e:
        print(file)
        print(e)

df = pd.DataFrame(records)

# Summary counts
summary = (
    df.groupby(["Term", "Phrase Found"])
      .size()
      .reset_index(name="Count")
      .sort_values(["Term", "Count"], ascending=[True, False])
)

output = Path(FOLDER) / "Terminology_Audit.xlsx"

with pd.ExcelWriter(output) as writer:

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    df.to_excel(
        writer,
        sheet_name="Occurrences",
        index=False
    )

print()
print("DONE")
print(output)
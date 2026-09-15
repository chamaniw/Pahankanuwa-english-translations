from docx import Document
from pathlib import Path
from collections import Counter
import re
import pandas as pd

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

before_counter = Counter()
after_counter = Counter()

for file in Path(FOLDER).rglob("*.docx"):

    try:
        doc = Document(file)

        text = " ".join(p.text for p in doc.paragraphs)

        words = re.findall(r"\b[\w'-]+\b", text)

        for i, word in enumerate(words):

            if word.lower() == "buddha":

                if i > 0:
                    before_counter[words[i-1]] += 1

                if i < len(words)-1:
                    after_counter[words[i+1]] += 1

    except Exception as e:
        print(f"Error: {file}")

before_df = pd.DataFrame(
    before_counter.most_common(),
    columns=["Word Before Buddha", "Count"]
)

after_df = pd.DataFrame(
    after_counter.most_common(),
    columns=["Word After Buddha", "Count"]
)

output = Path(FOLDER) / "Buddha_Word_Summary.xlsx"

with pd.ExcelWriter(output) as writer:
    before_df.to_excel(writer, sheet_name="Before Buddha", index=False)
    after_df.to_excel(writer, sheet_name="After Buddha", index=False)

print("Done")
print(output)
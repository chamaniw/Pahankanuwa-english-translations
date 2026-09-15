from docx import Document
from pathlib import Path
import re

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

father_buddha = 0
buddha_father = 0

for file in Path(FOLDER).rglob("*.docx"):

    try:
        doc = Document(file)

        text = " ".join(p.text for p in doc.paragraphs)

        father_buddha += len(
            re.findall(r"\bFather\s+Buddha\b", text, flags=re.IGNORECASE)
        )

        buddha_father += len(
            re.findall(r"\bBuddha\s+Father\b", text, flags=re.IGNORECASE)
        )

    except Exception:
        pass

print()
print("Father Buddha:", father_buddha)
print("Buddha Father:", buddha_father)
print("Total:", father_buddha + buddha_father)

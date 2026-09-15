from docx import Document
from pathlib import Path

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

count = 0

for file in Path(FOLDER).rglob("*.docx"):

    try:

        doc = Document(file)

        count += 1

    except Exception as e:

        print(file)
        print(e)

print("Files found:", count)
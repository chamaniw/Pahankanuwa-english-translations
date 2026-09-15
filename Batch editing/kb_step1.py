import pandas as pd

FOLDER = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations"

v3_file = FOLDER + r"\Nanananda_Dhamma_Index_V3.xlsx"
gemini_file = FOLDER + r"\Suttareference filled Gemini.xlsx"

print()
print("OPENING FILES...")
print()

# V3 workbook

v3 = pd.ExcelFile(v3_file)

print("V3 SHEETS")
for sheet in v3.sheet_names:
    print(" -", sheet)

print()

# Gemini workbook

gemini = pd.ExcelFile(gemini_file)

print("GEMINI SHEETS")
for sheet in gemini.sheet_names:
    print(" -", sheet)

print()

# Load first sheet

df = pd.read_excel(
    gemini_file,
    sheet_name=0
)

print("ROWS:", len(df))
print()

print("COLUMNS:")

for col in df.columns:
    print(" -", col)
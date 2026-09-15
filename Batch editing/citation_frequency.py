import pandas as pd

file = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations\V3_Step2.xlsx"

df = pd.read_excel(
    file,
    sheet_name="Pali_Citations_Clean"
)

freq = (
    df["Citation"]
    .value_counts()
    .reset_index()
)

freq.columns = [
    "Citation",
    "Count"
]

print()
print(freq.head(20))
import win32com.client as win32
from pathlib import Path
import pandas as pd

FOLDER = Path(r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations")

replacements = {
    "*Saṅkhāra* (Preparations / Formations / Volitional activity)": "*Saṅkhāra* (Formations)",
    "*Vedanā* (Feeling / Sensation)": "*Vedanā* (Feeling)",
    "*Taṇhā* (Craving / Thirst)": "*Taṇhā* (Craving)",
    "*Upādāna* (Grasping / Holding / Fuel)": "*Upādāna* (Grasping)",
    "*Bhava* (Becoming / Existence)": "*Bhava* (Becoming)",
    "*Jāti* (Birth / Arising)": "*Jāti* (Birth)",
    "*Jarāmaraṇa* (Decay and Death / Aging and Death)": "*Jarāmaraṇa* (Decay and Death)",
    "*Anattā* (Not-self / Non-self / Unsubstantial)": "*Anattā* (Not-self)",
    "*Anicca* (Impermanent / Inconstant)": "*Anicca* (Impermanent)",
    "*Dukkha* (Suffering / Unsatisfactoriness / Stress)": "*Dukkha* (Suffering)",
    "*Yoniso manasikāra* (Radical attention / Wise attention)": "*Yoniso manasikāra* (Wise attention)",
    "*Sakkāya-diṭṭhi* (Dogma of self-illusion / View of existing group/body)": "*Sakkāya-diṭṭhi* (View of existing group/body)",
    "*Anusaya* (Latency / Underlying tendency)": "*Anusaya* (Underlying tendency)",
    "*Asaṅkhata* (The Unprepared / Unconditioned / Uncompounded)": "*Asaṅkhata* (The Unconditioned)",
    "*Māyā* (Illusion / Magic trick)": "*Māyā* (Illusion)",
    "*Vipassanā* (Insight / Seeing through)": "*Vipassanā* (Insight)",
    "*Avijjā* (Ignorance / Nescience)": "*Avijjā* (Ignorance)",
    "*Paṭicca-samuppāda* (Law of Dependent Arising / Dependent Origination)": "*Paṭicca-samuppāda* (Law of Dependent Origination)",
    "*Kalyāṇamitta* (Good friend / Spiritual companion)": "*Kalyāṇamitta* (Spiritual companion)",
    "Nibbāna (The Mind Stilled / Extinction / Extinguishment)": "Nibbāna",
    "Nibbāna (The Mind Stilled / Extinction)": "Nibbāna"
}

word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False

report = []

for file in FOLDER.rglob("*.docx"):

    try:

        doc = word.Documents.Open(str(file))

        file_changes = 0

        for find_text, replace_text in replacements.items():

            result = doc.Content.Find.Execute(
                FindText=find_text,
                ReplaceWith=replace_text,
                Replace=2
            )

            if result:
                file_changes += 1

        if file_changes > 0:

            report.append({
                "File": file.name,
                "Replacement Types Applied": file_changes
            })

            print(f"Updated: {file.name}")

        doc.Save()
        doc.Close()

    except Exception as e:

        print(f"Error: {file.name}")
        print(e)

word.Quit()

pd.DataFrame(report).to_excel(
    FOLDER / "Terminology_Standardisation_Report.xlsx",
    index=False
)

print("DONE")
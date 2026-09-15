import win32com.client as win32
from pathlib import Path
import pandas as pd

FOLDER = Path(r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations")

replacements = {

    # ANATTĀ
    "*Anattā* (not-self/unsubstantial)":
    "*Anattā* (Not-self)",

    "*Anattā* [Not-self / Non-self / Unsubstantial]":
    "*Anattā* (Not-self)",

    "*Anattā* (not-self)":
    "*Anattā* (Not-self)",

    "*Anattā* (not-self/non-self/unsubstantial)":
    "*Anattā* (Not-self)",

    # VEDANĀ
    "*Vedanā* (feeling/sensation)":
    "*Vedanā* (Feeling)",

    "*Vedanā* [Feeling / Sensation]":
    "*Vedanā* (Feeling)",

    "*Vedanā* / feeling / sensation)":
    "*Vedanā* (Feeling)",

    "*Vedanā** (**Feeling / Sensation**)":
    "*Vedanā* (Feeling)",

    # SAṄKHĀRA
    "*Saṅkhāra* (preparations/formations)":
    "*Saṅkhāra* (Formations)",

    "*Saṅkhāra* (preparations/formations/volitional activity)":
    "*Saṅkhāra* (Formations)",

    "*Saṅkhāra* [preparations / formations / volitional activity]":
    "*Saṅkhāra* (Formations)",

    "*Saṅkhāra* [Preparations / Formations / Volitional activity]":
    "*Saṅkhāra* (Formations)",

    "*Saṅkhāra* / Formations / Volitional activity)":
    "*Saṅkhāra* (Formations)",

    # DUKKHA
    "*Dukkha* (suffering/unsatisfactoriness)":
    "*Dukkha* (Suffering)",

    "*Dukkha* [Suffering / Unsatisfactoriness / Stress]":
    "*Dukkha* (Suffering)",

    "*Dukkha* (suffering/unsatisfactoriness/stress)":
    "*Dukkha* (Suffering)",

    "*Dukkha** (**Suffering / Unsatisfactoriness / Stress**)":
    "*Dukkha* (Suffering)",

    "*Dukkha* (suffering)":
    "*Dukkha* (Suffering)",

    # YONISO MANASIKĀRA
    "*Yoniso manasikāra* [radical attention / wise attention]":
    "*Yoniso manasikāra* (Wise attention)",

    "*Yoniso manasikāra* (radical attention/wise attention)":
    "*Yoniso manasikāra* (Wise attention)",

    "*Yoniso manasikāra* (wise attention)":
    "*Yoniso manasikāra* (Wise attention)",

    # TAṆHĀ
    "*Taṇhā* [Craving / Thirst]":
    "*Taṇhā* (Craving)",

    "*Taṇhā* (craving/thirst)":
    "*Taṇhā* (Craving)",

    "*Taṇhā* (craving)":
    "*Taṇhā* (Craving)",

    # JĀTI
    "*Jāti* (birth/arising)":
    "*Jāti* (Birth)",

    "*Jāti* (arising/birth)":
    "*Jāti* (Birth)",

    # BHAVA
    "*Bhava* [Becoming / Existence]":
    "*Bhava* (Becoming)",

    "*Bhava* (becoming/existence)":
    "*Bhava* (Becoming)",

    "*Bhava* (becoming)":
    "*Bhava* (Becoming)",

    # ANICCA
    "*Anicca* [Impermanent / Inconstant]":
    "*Anicca* (Impermanent)",

    "*Anicca* (impermanence)":
    "*Anicca* (Impermanent)",

    "*Anicca* (impermanent/inconstant)":
    "*Anicca* (Impermanent)",

    "*Anicca* (impermanent)":
    "*Anicca* (Impermanent)",

    "*Anicca* (impermanence/inconstancy)":
    "*Anicca* (Impermanent)"
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

        if changes > 0:

            report.append({
                "File": file.name,
                "Replacement patterns found": changes
            })

            print(f"Updated: {file.name}")

        doc.Save()
        doc.Close()

    except Exception as e:

        print(f"Error: {file.name}")
        print(e)

word.Quit()

report_file = FOLDER / "Terminology_Cleanup_Pass3_Report.xlsx"

pd.DataFrame(report).to_excel(report_file, index=False)

print()
print("DONE")
print(f"Updated files: {len(report)}")
print(f"Report created: {report_file}")
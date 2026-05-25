import pandas as pd
import os
import pyperclip

def write_to_file(data, nazwa_pliku = "pomiary.xlsx"):
    df = pd.DataFrame([data], columns=range(1, len(data) + 1))

    writer = pd.ExcelWriter(nazwa_pliku, engine="xlsxwriter")

    df.to_excel(writer, sheet_name="Sheet1", index=False)

    pyperclip.copy("\t".join(str(x).replace(".", ",") for x in data))
    writer.close()
    os.startfile(nazwa_pliku)
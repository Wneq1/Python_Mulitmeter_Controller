# =============================================================================
# xls_writer.py — zapis wyników pomiarów do pliku Excel
#
# Zapisuje serię pomiarów jako jeden wiersz w arkuszu Sheet1.
# Jednocześnie kopiuje wartości do schowka (format z przecinkiem jako separatorem
# dziesiętnym — gotowe do wklejenia w polskim Excelu).
# =============================================================================

import pandas as pd
import os
import pyperclip


def write_to_file(data: list, nazwa_pliku: str = "pomiary.xlsx"):
    """
    Zapisuje listę wyników do pliku Excel i kopiuje do schowka.

    Parametry:
        data         — lista wartości float (wyniki pomiarów)
        nazwa_pliku  — ścieżka/nazwa pliku wyjściowego (domyślnie pomiary.xlsx)

    Po zapisaniu plik otwiera się automatycznie w domyślnej aplikacji.
    Schowek zawiera wartości oddzielone tabulatorami z przecinkiem jako
    separatorem dziesiętnym (kompatybilne z polskim Excel / LibreOffice Calc).
    """
    # Jeden wiersz, kolumny numerowane od 1
    df = pd.DataFrame([data], columns=range(1, len(data) + 1))

    writer = pd.ExcelWriter(nazwa_pliku, engine="xlsxwriter")
    df.to_excel(writer, sheet_name="Sheet1", index=False)
    writer.close()

    # Kopiuj do schowka: zamień kropkę na przecinek (polski format liczb)
    pyperclip.copy("\t".join(str(x).replace(".", ",") for x in data))

    # Otwórz plik w domyślnej aplikacji (Excel / LibreOffice)
    os.startfile(nazwa_pliku)

import pdfplumber
import openpyxl
from openpyxl import Workbook
from difflib import SequenceMatcher

if __name__ == '__main__':
    # Nom de la société
    path='../Fatture_Acquisto/Facture1.pdf'
    pdf = pdfplumber.open(path)
    page1=pdf.pages[0]
    IVA = page1.extract_tables()[0][0][0].find('IVA:')
    Deno = page1.extract_tables()[0][0][0].find('Denominazione')
    name = page1.extract_tables()[0][0][0][IVA+6:Deno-1]

    if len(name)>12:
        name=name[2:]
    print(name)
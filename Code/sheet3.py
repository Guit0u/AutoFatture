import pdfplumber
import openpyxl
from openpyxl import Workbook
from difflib import SequenceMatcher

if __name__ == '__main__':
    # Nom de la société
    path='../Fatture_Acquisto/Facture11.pdf'
    pdf = pdfplumber.open(path)
    page1=pdf.pages[0]
    Iden = page1.extract_tables()[0][0][0].find('IVA:')
    Deno = page1.extract_tables()[0][0][0].find('Denominazione')
    IVA = page1.extract_tables()[0][0][0][IVA+6:Deno]
    if len(IVA)>=12:
        IVA=IVA[1:13]

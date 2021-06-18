import pdfplumber
import openpyxl
from openpyxl import Workbook
from difflib import SequenceMatcher

if __name__ == '__main__':
    # Nom de la société
    path='../Fatture_Vendita/Fattura_13.PDF'
    pdf = pdfplumber.open(path)
    page1=pdf.pages[0]
    page1 = pdf.pages[0]
    # date
    x = page1.extract_text().find('Data')
    y = page1.extract_text().find('SEDE')
    date = page1.extract_text()[x + 5: y]

    # entreprise
    z = page1.extract_text().find('DESTINATARIO')
    t = page1.extract_text().find('Copia')
    name = page1.extract_text()[z + 12: t]

    #IVA
    IV = page1.extract_text().find('IVA')
    CF = page1.extract_text().find('C.F')
    IVA = page1.extract_text()[o+4: v]

    #NUM
    Nume = page1.extract_text().find('Numero')
    Data = page1.extract_text().find('Data')
    NUM = page1.extract_text()[num + 7: p]

    vii=page1.extract_text()

    print(vi)
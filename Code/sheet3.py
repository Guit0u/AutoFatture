import pdfplumber
import openpyxl
from openpyxl import Workbook
from difflib import SequenceMatcher

if __name__ == '__main__':
    path = '../Fatture_Vendita/Fattura_6.PDF'
    pdf = pdfplumber.open(path)

    # Produits

    Lines = []
    for page in pdf.pages:
        table = page.extract_tables({"vertical_strategy":"text"})
        print(table)
        for tables in table:
            for line in tables:
                print(type(line[4]))
                print(line[4])
                #if not 'ALIQUOTE' in line[1]:
                    #Lines.append(line)
    for i in range(len(Lines)):
        code = Lines[i][1]
        desc = Lines[i][2]
        quant = Lines[i][4].split(' ')
        quant = quant[-1]

        #print(code,desc,quant)
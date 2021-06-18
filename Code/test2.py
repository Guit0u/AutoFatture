import pdfplumber
import openpyxl
from openpyxl import Workbook

if __name__ == '__main__':
    path = 'Facture2.pdf'

    pdf = pdfplumber.open(path)
    #page=pdf.pages[0]
    #Deno = page.extract_text().find('Cod.')
    #Regime = page.extract_text().find('RIEPILOGHI')
    #name = page.extract_text()[Deno:Regime]
    for page in pdf.pages:
        x=page.extract_tables(table_settings={"horizontal_strategy":"text"})
        for lines in x:
            for line in lines:
                if (len(line)>7) and line[2]!='' and line[2]!='Quantità':
                    print(line)
    #print(x)

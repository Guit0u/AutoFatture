import pdfplumber
import openpyxl
from openpyxl import Workbook

if __name__ == '__main__':

    for _ in range(1,2):
        path = 'Facture'+str(_)+'.pdf'
    #path = 'Facture1.pdf'

        pdf = pdfplumber.open(path)
        #print(_)
        try: #Test si l'excel existe, dans ce cas là, on l'ouvre
            wb = openpyxl.load_workbook('test.xlsx')
            sheet1 = wb['Acquista']
            sheet2=wb['Vendita']
            sheet3 = wb['Inventario']

        except FileNotFoundError: #Sinon, on le crée
            wb = Workbook()
            sheet = wb.active
            sheet.title = 'Acquista'
            wb.create_sheet('Vendita')
            wb.create_sheet('Inventario')
            sheet1 = wb['Acquista']
            sheet2 = wb['Vendita']
            sheet3= wb['Inventario']

            sheet1.cell(1, 3).value='Cod Articolo'
            sheet1.cell(1, 4).value = 'Desczione'
            sheet1.cell(1, 5).value = 'Quantita'
            sheet1.cell(1, 6).value = 'Prezzo unitario'
            sheet1.cell(1, 7).value = 'UM'
            sheet1.cell(1, 8).value = 'Sconto o magg'
            sheet1.cell(1, 9).value = '%IVA'
            sheet1.cell(1, 10).value = 'Prezzo totale'
            sheet2.cell(1, 3).value = 'ARTICOLO'
            sheet2.cell(1, 4).value = 'DESCRIZIONE'
            sheet2.cell(1, 5).value = 'UNITÀ'
            sheet2.cell(1, 6).value = 'Q.TÀ'
            sheet2.cell(1, 7).value = 'IMPORTO U.'
            sheet2.cell(1, 8).value = 'IVA%SCONTO'
            sheet2.cell(1, 9).value = 'IMPORTO'
            sheet3.cell(1,3).value='Acquista'
            sheet3.cell(1,12).value='Vendita'



        max_r = sheet1.max_row #Donne l'emplacement pour écrire dans l'excel

        page1=pdf.pages[0]
        # Nom de la société
        Deno = page1.extract_tables()[0][0][0].find('Denominazione')
        Regime = page1.extract_tables()[0][0][0].find('Regime')
        name = page1.extract_tables()[0][0][0][Deno + 15:Regime-1]
        sheet1.cell(max_r + 2, 1).value = name
        # Date
        date = page1.extract_tables()[1][1][3]
        sheet1.cell(max_r + 2, 2).value = date

        # Numero documento
        numero = page1.extract_tables()[1][1][2]

        id= str(name + ' ' + date + ' ' + numero)
        fichier = open("datamm.txt", "r")
        lines=fichier.readlines()
        for line in lines:
            if str(line)==str(id)+"\n":
                print('Déjà rentrée')
                break
        fichier.close()
        fichier = open("datamm.txt", "a")
        fichier.write(id)
        fichier.write("\n")
        fichier.close()



        for page in pdf.pages:
            max_r = sheet1.max_row  # Donne l'emplacement pour écrire dans l'excel

            #Produits
            Lines=[]
            x = page.extract_tables(table_settings={"horizontal_strategy": "text"})
            for lines in x:
                for line in lines:
                    if (len(line) > 7) and line[2] != '' and line[2] != 'Quantità':
                        Lines.append(line)
                        #print(line)
                        #"""x = table[1] #a ameliorer
                    for i in range(len(Lines)):
                        #print(Lines[i])
                        #s = line[i]
                        #lis = s.split("\n")
                        for k in range(len(Lines[i])):
                            sheet1.cell(row=max_r+i + 2, column=k + 3).value=Lines[i][k]
                            sheet3.cell(row=max_r + i+1 , column=k + 3).value = Lines[i][k]

                        #if (s != ''):
                            #print(s)




            wb.save('test.xlsx')


        wb.close()
#"""

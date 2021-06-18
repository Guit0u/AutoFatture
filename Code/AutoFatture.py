import pdfplumber
import openpyxl
from openpyxl import Workbook
import sys
from PySide2 import QtCore, QtGui, QtWidgets
from pathlib import Path
import os
import sqlite3
from sqlite3 import Error

class MaFenetre(QtWidgets.QDialog):
    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent)

        self.boutonAchat = QtWidgets.QPushButton("Acquisto")
        self.boutonVente = QtWidgets.QPushButton("Venti")
        # Le champ de texte
        self.__champTexte = QtWidgets.QLineEdit("")
        self.labelMessage = QtWidgets.QLabel("")

        layout = QtWidgets.QGridLayout()
        layout.addWidget(self.__champTexte, 1, 1)
        layout.addWidget(self.labelMessage, 2, 1)
        layout.addWidget(self.boutonAchat, 3, 2)
        layout.addWidget(self.boutonVente, 3, 0)
        self.setLayout(layout)

        icone = QtGui.QIcon()
        icone.addPixmap(QtGui.QPixmap("bill.svg"))
        self.setWindowIcon(icone)

        self.boutonAchat.clicked.connect(self.genererAchat)
        self.boutonVente.clicked.connect(self.genererVente)

    def genererAchat(self):
        path = Path("../Fatture_Acquisto/" + self.__champTexte.text() + ".pdf")
        rep = os.getcwd()
        try:
            pdf = pdfplumber.open(path)

        except FileNotFoundError:
            print('file not found')
            self.__champTexte.clear()
            self.labelMessage.setText("This document doesn't exist")
            return
        try:  # Test si l'excel existe, dans ce cas là, on l'ouvre
            os.chdir(os.pardir)
            wb = openpyxl.load_workbook('Fatture.xlsx')
            sheet1 = wb['Acquista']
            sheet2 = wb['Vendita']
            sheet3 = wb['Inventario']

        except FileNotFoundError:  # Sinon, on le crée
            wb = Workbook()
            sheet = wb.active
            sheet.title = 'Acquista'
            wb.create_sheet('Vendita')
            wb.create_sheet('Inventario')
            sheet1 = wb['Acquista']
            sheet2 = wb['Vendita']
            sheet3 = wb['Inventario']
            sheet1.cell(1, 3).value = 'Cod Articolo'
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
            sheet3.cell(1, 3).value = 'Acquista'
            sheet3.cell(1, 12).value = 'Vendita'

        max_r = sheet1.max_row  # Donne l'emplacement pour écrire dans l'excel

        page1 = pdf.pages[0]
        # Nom de la société
        Deno = page1.extract_tables()[0][0][0].find('Denominazione')
        Regime = page1.extract_tables()[0][0][0].find('Regime')
        name = page1.extract_tables()[0][0][0][Deno + 15:Regime]
        sheet1.cell(max_r + 2, 1).value = name
        # Date
        date = page1.extract_tables()[1][1][3]
        sheet1.cell(max_r + 2, 2).value = date
        #numero de la commande
        # Numero documento
        numero = page1.extract_tables()[1][1][2]

        for page in pdf.pages:
            max_r = sheet1.max_row  # Donne l'emplacement pour écrire dans l'excel

            # Produits
            Lines = []
            x = page.extract_tables(table_settings={"horizontal_strategy": "text"})
            for lines in x:
                for line in lines:
                    if (len(line) > 7) and line[2] != '' and line[2] != 'Quantità':
                        Lines.append(line)
                        # print(line)
                        # """x = table[1] #a ameliorer
                    for i in range(len(Lines)):
                        for k in range(len(Lines[i])):
                            sheet1.cell(row=max_r + i + 2, column=k + 3).value = Lines[i][k]
                            sheet3.cell(row=max_r + i+1, column=k + 3).value = Lines[i][k]
        F=str(name)
        D=str(date)
        N=str(numero)
        if check(F,D,N):
            self.labelMessage.setText("This bill has already been registered")
            self.__champTexte.clear()
            return

        insert_Fatture='''INSERT INTO Fatture(Fournisseur, Date,NumCom )
                            VALUES(?,?,?)'''
        tuple=(F,D,N)
        cur.execute(insert_Fatture,tuple)
        conn.commit()

        wb.save('Fatture.xlsx')
        wb.close()
        os.chdir(rep)
        print("success")
        self.labelMessage.setText("Success")
        self.__champTexte.clear()

    def genererVente(self):
        path = Path("../Fatture_Vendita/" + self.__champTexte.text() + ".pdf")
        rep = os.getcwd()
        try:
            pdf = pdfplumber.open(path)

        except FileNotFoundError:
            print('file not found')
            self.__champTexte.clear()
            self.labelMessage.setText("This document doesn't exist")
            return
        try:  # Test si l'excel existe, dans ce cas là, on l'ouvre
            os.chdir(os.pardir)
            wb = openpyxl.load_workbook('Fatture.xlsx')
            sheet1 = wb['Acquista']
            sheet2 = wb['Vendita']
            sheet3 = wb['Inventario']

        except FileNotFoundError:  # Sinon, on le crée
            wb = Workbook()
            sheet = wb.active
            sheet.title = 'Acquista'
            wb.create_sheet('Vendita')
            wb.create_sheet('Inventario')
            sheet1 = wb['Acquista']
            sheet2 = wb['Vendita']
            sheet3 = wb['Inventario']
            sheet1.cell(1, 3).value = 'Cod Articolo'
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
            sheet3.cell(1, 3).value = 'Acquista'
            sheet3.cell(1, 12).value = 'Vendita'

        max_r = sheet2.max_row  # Donne l'emplacement pour écrire dans l'excel

        page1 = pdf.pages[0]
        # date
        x = page1.extract_text().find('Data')
        y = page1.extract_text().find('SEDE')
        date = page1.extract_text()[x + 5: y]
        sheet2.cell(max_r + 2, 2).value = date
        # entreprise
        z = page1.extract_text().find('DESTINATARIO')
        t = page1.extract_text().find('Copia')
        name = page1.extract_text()[z + 12: t]
        sheet2.cell(max_r + 2, 1).value = name

        # Produits
        max_r = sheet2.max_row  # Donne l'emplacement pour écrire dans l'excel
        Lines = []
        for page in pdf.pages:
            table = page.extract_tables(table_settings={"vertical_strategy": "text"})
            for tables in table:
                for line in tables:
                    if not 'ALIQUOTE' in line[1]:
                        Lines.append(line)
                    for i in range(len(Lines)):
                        # print(Lines[i])
                        # s = line[i]
                        # lis = s.split("\n")
                        for k in range(len(Lines[i])):
                            sheet2.cell(row=max_r + i + 2, column=k + 2).value = Lines[i][k]
                            sheet3.cell(row=max_r + i + 1, column=k + 11).value = Lines[i][k]
                            # print(c)
        wb.save('Fatture.xlsx')
        wb.close()
        os.chdir(rep)
        print("success")
        self.labelMessage.setText("Success")
        self.__champTexte.clear()




def check(Fournisseur,Date, NumCom):
    b=(str(Date),str(NumCom),str(Fournisseur))
    print(b)
    check="""SELECT Date, NumCom, Fournisseur
    FROM Fatture as F
    WHERE F.Date= ? AND F.NumCom = ? AND F.Fournisseur = ?"""
    cur.execute(check,b)
    boo = ''
    for i in cur:
        print(i)
        boo=i
    if(boo!=''):
        return True
    return False





conn = None
rep = os.getcwd()
os.chdir(os.pardir)
dbf = str(Path("DB/database.db").absolute())
os.chdir(rep)
try:
    conn = sqlite3.connect(dbf)
    print(sqlite3.version)

    cur = conn.cursor()
    table_factures = '''CREATE TABLE IF NOT EXISTS Fatture(
                              Fournisseur TEXT,
                              Date TEXT,
                              NumCom TEXT
                             )'''
    cur.execute(table_factures)
    table_fournisseurs = '''CREATE TABLE IF NOT EXISTS Fournisseurs(
                                Nom TEXT,
                                IVA INT
                                )'''
    cur.execute(table_fournisseurs)



    for row in cur:
        print(row)
    conn.commit()

    app = QtWidgets.QApplication(sys.argv)
    dialog = MaFenetre()
    dialog.exec_()

except Error as e:
    print(e)
finally:
    if conn:
       conn.close()

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

        #les boutons

        self.boutonAchat = QtWidgets.QPushButton("Acquisto")
        self.boutonVente = QtWidgets.QPushButton("Venti")

        self.boutonAddClient = QtWidgets.QPushButton("Add client")

        # Les champs de texte
        self.__champTexte = QtWidgets.QLineEdit("")
        self.labelMessage = QtWidgets.QLabel("")

        self.labelAdd = QtWidgets.QLabel("")
        self.__champIva = QtWidgets.QLineEdit("IVA")
        self.__champNom = QtWidgets.QLineEdit("Nom")

        layout = QtWidgets.QGridLayout()
        layout.addWidget(self.__champTexte, 1, 1)
        layout.addWidget(self.labelMessage, 2, 1)
        layout.addWidget(self.boutonAchat, 3, 2)
        layout.addWidget(self.boutonVente, 3, 0)
        layout.addWidget(self.__champIva, 5, 0)
        layout.addWidget(self.__champNom, 5, 2)
        layout.addWidget(self.boutonAddClient,7,1)
        layout.addWidget(self.labelAdd,6,1)
        self.setLayout(layout)

        icone = QtGui.QIcon()
        icone.addPixmap(QtGui.QPixmap("bill.svg"))
        self.setWindowIcon(icone)

        self.boutonAchat.clicked.connect(self.genererAchat)
        self.boutonVente.clicked.connect(self.genererVente)
        self.boutonAddClient.clicked.connect(self.AddClientBouton)


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
            os.chdir(rep)

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
        name = page1.extract_tables()[0][0][0][Deno + 15:Regime-1]

        # IVA De la société
        Iden = page1.extract_tables()[0][0][0].find('IVA:')
        Deno = page1.extract_tables()[0][0][0].find('Denominazione')
        IVA = page1.extract_tables()[0][0][0][Iden + 6:Deno-1]
        if len(IVA) >= 12:
            IVA = IVA[1:13]
        IVA=IVA.split('\n')[0]
        # Date de la commande
        Date = page1.extract_tables()[1][1][3]
        
        # Numero de la commande
        NumCom = page1.extract_tables()[1][1][2]

        I=str(IVA)
        D=str(Date)
        N=str(NumCom)

        if checkA(I,D,N):
            self.labelMessage.setText("This bill has already been registered")
            self.__champTexte.clear()
            print(I+D+N)
            wb.save('Fatture.xlsx')
            wb.close()
            os.chdir(rep)
            return

        if checkFourn(I):
            insert_fournisseur= '''INSERT INTO Fournisseurs(Nom,IVA)
                                        VALUES(?,?)'''
            tuple = (name, I)
            cur.execute(insert_fournisseur, tuple)
            print('nouveau client!')
            conn.commit()

        sheet1.cell(max_r + 2, 1).value = name
        sheet1.cell(max_r + 2, 2).value = Date

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

        insert_Fatture='''INSERT INTO FattureA(IVA, Date,NumCom )
                            VALUES(?,?,?)'''
        tuple=(I,D,N)
        cur.execute(insert_Fatture,tuple)
        conn.commit()

        wb.save('Fatture.xlsx')
        wb.close()
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
            os.chdir(rep)

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
        dataa = page1.extract_text().find('Data')
        sede = page1.extract_text().find('SEDE')
        Date = page1.extract_text()[dataa + 5: sede-1]


        # entreprise
        dest = page1.extract_text().find('DESTINATARIO')
        copia = page1.extract_text().find('Copia')
        name = page1.extract_text()[dest + 13: copia-1]

        # IVA
        IV = page1.extract_text().find('IVA')
        CF = page1.extract_text().find('C.F')
        IVA = page1.extract_text()[IV + 4: CF-1]

        # NUM
        Nume = page1.extract_text().find('Numero')
        Data = page1.extract_text().find('Data')
        NumCom = page1.extract_text()[Nume + 7: Data-1]

        I = str(IVA)
        D = str(Date)
        N = str(NumCom)
        if checkB(I, D, N):
            self.labelMessage.setText("This bill has already been registered")
            self.__champTexte.clear()
            print(I +D +N)
            return

        if checkFourn(I):
            insert_fournisseur = '''INSERT INTO Fournisseurs(Nom,IVA)
                                        VALUES(?,?)'''
            tuple = (name, I)
            cur.execute(insert_fournisseur, tuple)
            print('nouveau client!')
            conn.commit()

        sheet1.cell(max_r + 2, 1).value = name
        sheet1.cell(max_r + 2, 2).value = Date
        sheet2.cell(max_r + 2, 2).value = Date
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


        insert_Fatture = '''INSERT INTO FattureV(IVA, Date,NumCom )
                                    VALUES(?,?,?)'''
        tuple = (I, D, N)
        cur.execute(insert_Fatture, tuple)
        conn.commit()

        wb.save('Fatture.xlsx')
        wb.close()

        print("success")
        self.labelMessage.setText("Success")
        self.__champTexte.clear()

    def AddClientBouton(self):
        IVA = self.__champIva.text()
        Nom = self.__champNom.text()
        if len(IVA) > 11:
            print("IVA too long")
            self.labelAdd.setText("IVA too long")
            self.__champIva.clear()
            return
        elif len(IVA) < 11:
            self.labelAdd.setText("IVA too short")
            self.__champIva.clear()
            return
        else:

           if addClient(IVA,Nom):
               self.labelAdd.setText("client succesfully added")
           else:
               self.labelAdd.setText("client already exists")

           self.__champIva.clear()
           self.__champNom.clear()




def checkA(IVA,Date,NumCom):
    tuple=(str(IVA),str(Date),str(NumCom))
    check="""SELECT IVA,Date, NumCom
    FROM FattureA as F
    WHERE F.IVA = ? AND F.Date = ? AND F.NumCom = ?"""
    cur.execute(check,tuple)
    boo = ''
    for i in cur:
        print(i)
        boo=i
    if(boo==''):
        return False
    return True

def checkB(IVA,Date, NumCom):
    tuple=(str(IVA),str(Date),str(NumCom))
    check="""SELECT IVA,Date, NumCom
    FROM FattureV as F
    WHERE F.IVA = ? AND F.Date = ? AND F.NumCom = ?"""
    cur.execute(check,tuple)
    boo = ''
    for i in cur:
        print(i)
        boo=i
    if(boo==''):
        return False
    return True


def checkFourn(IVA):
    tuple = (str(IVA),)
    check = """SELECT IVA
        FROM Fournisseurs as F
        WHERE F.IVA= ?  """
    cur.execute(check, tuple)
    bool=''
    for row in cur:
        bool = row
    if bool=='':
        return True
    return False

def addClient(IVA,Nom):
    if(checkFourn(IVA)):
        tuple = (str(IVA),str(Nom))
        requestAdd = """INSERT INTO Fournisseurs(IVA,Nom)
                            VALUES(?,?)"""
        cur.execute(requestAdd,tuple)
        conn.commit()
        return True
    else:
        print("Deja client")
        return False

def SuppClient(IVA):
    pass

conn = None
rep = os.getcwd()
os.chdir(os.pardir)
dbf = str(Path("DB/database.db").absolute())
os.chdir(rep)
try:
    conn = sqlite3.connect(dbf)
    print(sqlite3.version)

    cur = conn.cursor()
    table_facturesA = '''CREATE TABLE IF NOT EXISTS FattureA(
                              IVA TEXT,
                              Date TEXT,
                              NumCom TEXT
                             )'''
    cur.execute(table_facturesA)
    table_facturesV = '''CREATE TABLE IF NOT EXISTS FattureV(
                                  IVA TEXT,
                                  Date TEXT,
                                  NumCom TEXT
                                 )'''
    cur.execute(table_facturesV)
    table_fournisseurs = '''CREATE TABLE IF NOT EXISTS Fournisseurs(
                                Nom TEXT,
                                IVA TEXT
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

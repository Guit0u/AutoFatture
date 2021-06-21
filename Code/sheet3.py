import pdfplumber
import openpyxl
from openpyxl import Workbook
from difflib import SequenceMatcher

if __name__ == '__main__':
    str='84,25'
    #x=int(str,10)
    x=float(str.strip().split(" ")[0].replace(',', '.'))
    y=float(x)
    print(y)
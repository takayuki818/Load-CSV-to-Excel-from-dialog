import openpyxl
import csv
import tkinter as tk
from tkinter import filedialog
import sys
def パス取得():
    root=tk.Tk()
    root.withdraw()
    file_path=filedialog.askopenfilename()
    return file_path
def CSV読込():
    file_path=パス取得()
    if file_path=='':
        print('読込CSVファイルがありません')
        sys.exit()
    output_excel_path='CSV展開.xlsx'
    encoding='utf-8'
    delimiter=','
    newline=''
    header_rows=1
    wb=openpyxl.Workbook()
    sh=wb.active
    with open(file_path,'r',encoding=encoding,newline=newline) as csv_file:
        csv_reader=csv.reader(csv_file,delimiter=delimiter)
        for _ in range(header_rows):
            next(csv_reader)
        for row in csv_reader:
            sh.append(row)
    wb.save(output_excel_path)
    wb.close
CSV読込()
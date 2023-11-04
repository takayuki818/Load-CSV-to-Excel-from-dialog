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
def CSV読込(file_path):
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
        for row_idx,row in enumerate(csv_reader):
            if row_idx<header_rows:
                continue
            for col_idx,value in enumerate(row):
                sh.cell(row=row_idx-header_rows+1,column=col_idx+1,value=value)
    wb.save(output_excel_path)
    wb.close
CSV読込(パス取得())

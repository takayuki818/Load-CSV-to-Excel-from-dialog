import openpyxl
import csv
import PySimpleGUI as sg
from tkinter import filedialog
import sys
def CSV読込(file_path,header_rows):
    if file_path==''or file_path==None:
        sg.popup('読込CSVファイルがありません')
        sys.exit()
    if not isinstance(header_rows,int):
        sg.popup('見出し行数が整数ではありません')
        sys.exit()
    output_excel_path='CSV展開.xlsx'
    encoding='utf-8'
    delimiter=','
    newline=''
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
layout=[[sg.Text('展開するCSVファイルのパスを入力（Browseから選択）')], 
        [sg.InputText(),sg.FileBrowse(key='file_path', file_types=(('CSVファイル', '*.csv'),))],
        [sg.Text('読込から除外する見出し行数を入力（全行読込の場合は「0」）'),sg.InputText(key='header_rows',size=(5,))],
        [sg.Button('CSV読込実行',key='OK')]]
window=sg.Window('CSV読込設定',layout)
while True:
    event,values=window.read()
    if event==sg.WIN_CLOSED:
        break
    elif event=='OK':
        file_path=values['file_path']
        header_rows=int(values['header_rows'])
        CSV読込(file_path,header_rows)
        sg.popup('CSVファイル読込が完了しました')
        break
window.close()
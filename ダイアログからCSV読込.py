import openpyxl
import csv
import PySimpleGUI as sg
def CSV読込(file_path,header_rows,encoding,delimiter):
    output_excel_path='CSV展開.xlsx'
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
def 整数判定(value):
    try:
        int(value, 10) 
    except ValueError:
        return False
    else:
        return True
def main():
    layout=[[sg.Text('展開するCSVファイルのパスを入力（Browseから選択）')], 
            [sg.InputText(),sg.FileBrowse(key='file_path', file_types=(('CSVファイル', '*.csv'),))],
            [sg.Text('読込除外見出行数（全行読込の場合は「0」）'),sg.InputText(key='header_rows',size=(5,))],
            [sg.Text('文字コードを選択'),sg.Combo(['UTF-8','SHIFT-JIS','CP932'],default_value='UTF-8',size=(6,1), key='encoding'),
             sg.Text('区切り文字を入力'),sg.InputText(',',justification='center',size=(5,),key='delimiter')],
            [sg.Button('CSV読込実行',key='OK')]]
    window=sg.Window('CSV読込設定',layout)
    while True:
        event,values=window.read()
        if event==sg.WIN_CLOSED:
            break
        elif event=='OK':
            file_path=values['file_path']
            header_rows=values['header_rows']
            encoding=values['encoding']
            delimiter=values['delimiter']
            if 整数判定(header_rows):
                if file_path=='':
                    sg.popup('CSVファイルのパスを入力してください')
                else:
                    header_rows=int(values['header_rows'])
                    CSV読込(file_path,header_rows,encoding,delimiter)
                    sg.popup('CSVファイル読込が完了しました')
                    break
            else:
                sg.popup('見出し行数には整数を入力してください')
    window.close()
if __name__ == '__main__':
    main()


# 差込データのファイル名
mailmerge_xlsx_file_name = "差込データ.xlsx"
# 分割した差込データのファイルを保存するフォルダ名
xlsx_folder_name = "分割した差込データ"

import os
import shutil
import win32com.client

# 分割した差込データのファイルを保存するディレクトリをリセットする
xlsx_folder = os.path.abspath(xlsx_folder_name)
if os.path.exists(xlsx_folder):
    shutil.rmtree(xlsx_folder)
os.mkdir(xlsx_folder)


xlsx = win32com.client.Dispatch("Excel.Application")

# ファイル名を直接指定するとWorkbooks.Openでエラーになる
xlsx_path = os.path.abspath(mailmerge_xlsx_file_name)
# 差込データのファイルはプログラムと同じフォルダにあるとする
data_book = xlsx.Workbooks.Open(xlsx_path, ReadOnly="True")

# 差込データは最初のシートに保存されているとする
sheet = data_book.Worksheets(1)

c = 1
while 1:
    value = sheet.cells(1, c)
    if value.value is None:
        break
    c = c+1

max_column = c

r = 2
while 1:
    value = sheet.Cells(r, 1)

    # データが書かれていない行はA桁が空白とする
    # A桁が空白の行ならば処理を終了
    if value.value is None:
        break

    name = value.value
    print(name)
    divided_xlsx_path = os.path.join(xlsx_folder,(name+".xlsx"))

    divide_book = xlsx.Workbooks.Add()
    divide_sheet = divide_book.Worksheets(1)
    for c in range(1, max_column):
        divide_sheet.Cells(1, c).value = sheet.Cells(1, c).value
        divide_sheet.Cells(2, c).value = sheet.Cells(r, c).value

    divide_book.SaveAs(divided_xlsx_path)
    divide_book.Close()

    r = r+1

data_book.Close()

xlsx.Quit()

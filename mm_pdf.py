# 分割した差込データのファイルを保存するフォルダ名
xlsx_folder_name = "分割した差込データ"
# 作成したファイルを保存するフォルダ名
pdf_folder_name = "作成したPDFファイル"
docx_folder_name = "作成したWordファイル"

docx_file_name = "差込印刷 元文書.docx"

import os
import shutil
import win32com.client

docx = win32com.client.Dispatch("Word.Application")

docx_path = os.path.abspath(docx_file_name)
document = docx.Documents.Open(docx_path)

# 作成したPDFファイルを保存するディレクトリをリセットする
pdf_folder = os.path.abspath(pdf_folder_name)
if os.path.exists(pdf_folder):
    shutil.rmtree(pdf_folder)
os.mkdir(pdf_folder)

# 作成したWordファイルを保存するディレクトリをリセットする
docx_folder = os.path.abspath(docx_folder_name)
if os.path.exists(docx_folder):
    shutil.rmtree(docx_folder)
os.mkdir(docx_folder)

for xlsx_file in os.listdir(xlsx_folder_name):
    print(xlsx_file)

    name, ext = os.path.splitext(xlsx_file)
    xlsx_path = os.path.join(os.getcwd(), xlsx_folder_name, xlsx_file)
    pdf_file = os.path.join(os.getcwd(), pdf_folder_name, (name+".pdf"))
    docx_file = os.path.join(os.getcwd(), docx_folder_name, (name+".docx"))

    document.MailMerge.OpenDataSource(xlsx_path, SQLStatement="SELECT * FROM [Sheet1$]")
    document.MailMerge.Destination = win32com.client.constants.wdSendToNewDocument
    document.MailMerge.Execute()

    document_pdf = docx.Documents(1)
    document_pdf.SaveAs(docx_file)
    document_pdf.SaveAs(pdf_file, win32com.client.constants.wdFormatPDF)
    document_pdf.Close()

document.Close()
docx.Quit()

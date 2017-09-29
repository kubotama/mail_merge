# mail_merge
# 差し込み印刷の結果を別ファイルに分割

仕事で一部だけを差し替えて文書を作る場合、Wordの差し込み印刷が便利なんだけど、差し込んだ結果は一つのファイルにまとめられてしまうのが仕様らしいです。たとえば招待状とかで、差し出し相手ごとに氏名とか所属とかを差し込んだ文書を作るとき、それぞれ別のファイルになってくれるとうれしいんだけど、どうやらそういう機能はないらしいです。ないんだったら、と作ってしまいました。

考え方としては以下の通りです。Excelとかで作った差し込むデータが入っているファイルごとにファイルが出力されるみたいなので、差し込むデータのファイルを分割すれば出力先のファイルも分割されます。なので、以下の2つの機能を実行するプログラムをPythonで作ってみました。

1. 差し込むデータの入っているファイルを1行ずつに分割する。
1. 分割したファイルごとに差し込み印刷を実行する。

# 前提条件
以下の環境で作成および動作を検証しています。

- Microsoft Windows10 Pro 64bit
- Microsoft Office 365 Solo
- Python 3.6.2 64bit (グローバルに利用可能としてインストール)
- [Python for Windows Extensions(pywin32)](https://sourceforge.net/projects/pywin32/) [Build 221](https://sourceforge.net/projects/pywin32/files/pywin32/Build%20221/)

## 環境の導入について

pywin32の導入および利用について以下の2点で注意が必要でした。pywin32のドキュメントなどを見つけることができていないため、これが正しいのかは不明です。

### 初回のimport

原因は未確認ですが、Python for Windows Extensionsをimportしようとしてエラーが発生しましたが、以下の手順で回避できました。

1. 管理者モードでコマンドプロンプトを起動します。
1. PythonのCLIモードに入ります。
1. import win32comおよびimport win32com.clientを実行します。

一回importをすると、それ以降は管理者モードでなくてもimport可能です。

### 定数の参照

Microsoft Officeのライブラリで定義されている定数(wdSendToNewDocumentなど)を利用する場合には、あらかじめ以下の手順が必要のようです。

1. 管理者モードでPythonWinを起動します。Pythonをグローバルではなく個人のみが利用可能として導入している場合には違うかもしれません。
1. ToolsメニューからCOM Makepy utilityを選択します。
1. Type LibraryからMicrosoft Word 16.0 Object Library(8.7)を選択してOKボタンをクリックする。数字はインストールされているOfficeのバージョンとかで変わるみたいなので、適当に調整してください。PythonWinを一般ユーザーで起動していると、ここでPermission Deniedだと表示されます。

Interactive Windowに以下のように表示されれば完了です。

```
Genrating to C:\Program Files\Python36\lib\site-packages\win32com\gen_py\000XXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXXXXXXXXX.py
Building definitions from type library...
Generating...
Importing module
```

これで定数はwin32com.client.constants.wdSendToNewDocuementとして参照可能になります。


# 差し込みデータのファイルを分割

[xlsx_divide.py](https://github.com/kubotama/mail_merge/blob/master/xlsx_divide.py)が処理します。このプログラムは、差し込みデータが保存されているExcelファイルを分割します。

## 差し込みデータのファイル形式

差し込みデータが保存されているExcelファイルの形式は以下の通りとします。

- 差し込みデータは最初のシートに保存されています。
- 1行目がヘッダで2行目以降にデータが保存されています。
- ヘッダが空白でない列まで有効なデータが保存されています。
- A列が空白でない行まで有効なデータが保存されています。

サンプルの差込データ.xlsxの最初のシート(Sheet1)は以下の通りです。

| 氏名 | 所属 |
|:-:|:-:|
| 柴田勲 | センター |
| 土井正三 | セカンド |
| 王貞治  | ファースト |
| 長嶋茂雄 | サード |
| 末次民夫 | ライト |
| 高田繁 | レフト |
| 黒江透修 | ショート |
| 森昌彦 | キャッチャー |
| 堀内恒夫 | ピッチャー |

## 参照するファイルやフォルダの指定

差し込みデータが保存されているファイルは以下の行で指定しています。

> ```python
> mailmerge_xlsx_file_name = "差込データ.xlsx"
> ```

分割されたファイルは1行目がヘッダ行、2行目が差し込みデータで2行しかないファイルとして作成されて、以下の行で指定したフォルダにA桁のデータをファイル名として保存されます。

> ```python
> xlsx_folder_name = "分割した差込データ"
> ```

## 以前に作成されたファイルのクリア

以下では、以前に分割したファイルが残っていれば削除しています。

> ```python
> xlsx_folder = os.path.abspath(xlsx_folder_name)
> if os.path.exists(xlsx_folder):
>    shutil.rmtree(xlsx_folder)
> os.mkdir(xlsx_folder)
> ```

## データ・ファイルのオープン

以下でExcelを起動して、差し込みデータが保存されているExcelファイルをオープンしています。

> ```python
> xlsx = win32com.client.Dispatch("Excel.Application")
> # ファイル名を直接指定するとWorkbooks.Openでエラーになる
> xlsx_path = os.path.abspath(mailmerge_xlsx_file_name)
> # 差込データのファイルはプログラムと同じフォルダにあるとする
> data_book = xlsx.Workbooks.Open(xlsx_path, ReadOnly="True")
> ```

差し込みデータは指定されたExcelファイルの最初のシートに保存されている前提です。もし最初以外のシートの場合には、以下を適当に修正してください。

> ```python
> # 差込データは最初のシートに保存されているとする
> sheet = data_book.Worksheets(1)
> ```

## 差し込みデータの分割

分割したファイルには、ヘッダ行が空白でない列まで各行のデータがコピーされます。以下ではヘッダがどの列まで空白でないのかを調べています。セルの値はvalue属性で参照できます。

> ```python
> c = 1
> while 1:
> if sheet.cells(1, c).value is None:
>        break
>    c = c+1
>
> max_column = c
> ```

1行目はヘッダなので、2行目から処理します。A列が空白の行があれば、そこで処理を終了します。

> ```python
> r = 2
>while 1:
>    # データが書かれていない行はA列が空白とする
>    # A列が空白の行ならば処理を終了
>    name = sheet.Cells(r, 1).value
>    if name is None:
>        break
> ```

分割したた差し込みデータを保存するファイルを作成して、作成したファイルの最初のシートにヘッダとデータをコピーします。

> ```python
> divide_book = xlsx.Workbooks.Add()
> divide_sheet = divide_book.Worksheets(1)
> for c in range(1, max_column):
>     divide_sheet.Cells(1, c).value = sheet.Cells(1, c).value
>     divide_sheet.Cells(2, c).value = sheet.Cells(r, c).value
> ```

A列のデータから作成したファイル名で保存します。

> ```python
>     divided_xlsx_path = os.path.join(xlsx_folder,(name+".xlsx"))
>     divide_book.SaveAs(divided_xlsx_path)
>     divide_book.Close()
> ```

これで指定したフォルダに分割された差し込みデータが作成されました。

# 分割したデータごとに差し込み印刷

[mm_pdf.py](https://github.com/kubotama/mail_merge/blob/master/mm_pdf.py)が処理します。このプログラムは、指定されたフォルダに保存されているExcelファイルで差し込み印刷を実行して、WordとPDFのそれぞれのファイル形式で保存します。

## 参照するファイルやフォルダの指定

差し込み印刷の元文書をWordで作成して、データソースとして分割する前のExcelファイルを指定しておいてください。以下では、分割された差し込みデータのExcelファイルが保存されているフォルダ、差し込み印刷で作成したWordおよびPDFのそれぞれの形式で保存するフォルダ、差し込み印刷の元文書とするWordファイルを指定しています。

> ```python
> # 分割した差込データのファイルを保存するフォルダ名
> xlsx_folder_name = "分割した差込データ"
># 作成したファイルを保存するフォルダ名
> pdf_folder_name = "作成したPDFファイル"
> docx_folder_name = "作成したWordファイル"
>
> docx_file_name = "差込印刷 元文書.docx
> ```

## 以前に作成されたファイルのクリア

以下では、以前に分割したファイルが残っていれば削除しています。

> ```python
> # 作成したPDFファイルを保存するディレクトリをリセットする
>pdf_folder = os.path.abspath(pdf_folder_name)
>if os.path.exists(pdf_folder):
>    shutil.rmtree(pdf_folder)
>os.mkdir(pdf_folder)
>
> # 作成したWordファイルを保存するディレクトリをリセットする
> docx_folder = os.path.abspath(docx_folder_name)
> if os.path.exists(docx_folder):
>    shutil.rmtree(docx_folder)
> os.mkdir(docx_folder)
> ```

## 差し込み印刷の元文書のオープン

以下でWordを起動して、差し込み印刷の元文書となるWordファイルをオープンしています。

> ```python
> docx = win32com.client.Dispatch("Word.Application")
> docx_path = os.path.abspath(docx_file_name)
> document = docx.Documents.Open(docx_path)
> ```

## 差し込み印刷の実行

分割された差し込みデータのファイル名から、作成するWord, PDF形式のファイル名を作成しています。

> ```python
>   name, ext = os.path.splitext(xlsx_file)
>   xlsx_path = os.path.join(os.getcwd(), xlsx_folder_name, xlsx_file)
>   pdf_file = os.path.join(os.getcwd(), pdf_folder_name, (name+".pdf"))
>   docx_file = os.path.join(os.getcwd(), docx_folder_name, (name+".docx"))
> ```

差し込み印刷をするときのデータソースと差し込み印刷の結果を別ファイルとする指定をしてから差し込み印刷を実行します。

> ```python
>   document.MailMerge.OpenDataSource(xlsx_path, SQLStatement="SELECT * FROM [Sheet1$]")
>   document.MailMerge.Destination = win32com.client.constants.wdSendToNewDocument
>   document.MailMerge.Execute()
> ```

差し込み印刷の結果をWordおよびPDF形式で保存します。

> ```python
>   document_pdf = docx.Documents(1)
>   document_pdf.SaveAs(docx_file)
>   document_pdf.SaveAs(pdf_file, win32com.client.constants.wdFormatPDF)
>   document_pdf.Close()
> ```

これで差し込み印刷の結果が別ファイルに分割されました。

# 今後の予定

今後の予定としては、以下の機能追加が可能かを調べてみたいと考えています。

- [ ] withブロックに対応して、明示的なQuitやCloseを不要とする。
- [ ] 行単位の読み込みをイテレーションとする。

Pythonの修行中なので、もっとPythonらしい書き方があれば教えてもらえると有難いです。

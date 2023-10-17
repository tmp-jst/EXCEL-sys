import os
import sys
import tkinter.filedialog
from tkinter import messagebox
import tkinter as tk
import pandas as pd
import string

# ダイアログを表示し、ファイル選択
root = tk.Tk()
root.withdraw()
fTyp = [("", "*.xlsx")]
iDir = os.path.abspath(os.path.dirname(__file__))
ret = messagebox.askyesno("入出力データの指定", "エクセルファイルを選択してください")
if not ret:
    messagebox.showerror("プログラム終了", "ファイルが選択されませんでした\nプログラムを終了します")
    sys.exit()

input_fname = tkinter.filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
if not input_fname:
    messagebox.showerror("プログラム終了", "ファイルが選択されませんでした\nプログラムを終了します")
    sys.exit()

# エクセルファイルを読み込む
checklist_data = pd.read_excel(input_fname, header=0)

# カラム選択用のダイアログを表示
def show_column_selection_dialog():
    column_selection_dialog = tk.Toplevel(root)
    column_selection_dialog.title("カラムの選択")

    # カラムヘッダーを読み込み（ヘッダー行から取得することを想定）
    # この例では仮のヘッダーを使用
    column_headers = checklist_data.columns

    columns = []
    for col_header in column_headers:
        col_var = tk.IntVar()
        columns.append(col_var)
        chk = tk.Checkbutton(column_selection_dialog, text=col_header, variable=col_var)
        chk.pack()

    btn_select_columns = tk.Button(column_selection_dialog, text="選択したカラムを表示", command=select_columns)
    btn_select_columns.pack()

# カラムの選択用ダイアログから選択されたカラムを処理
selected_columns = []

def select_columns():
    for col_var, col_header in zip(columns, column_headers):
        if col_var.get():
            selected_columns.append(col_header)
    if not selected_columns:
        messagebox.showerror("エラー", "少なくとも1つのカラムを選択してください")
    else:
        generate_html(selected_columns)

def generate_html(selected_columns):
    # 選択されたカラムのデータのみを抽出
    selected_data = checklist_data[selected_columns]

    # ここで選択されたデータをテンプレートのHTMLに結合（この部分を実際の結合方法に合わせてカスタマイズ）
    html_body = ""
    for index, row in selected_data.iterrows():
        for col in selected_columns:
            html_body += f"<p><b>{col}:</b> {row[col]}</p>"

    # テンプレート結合
    template_txt_html = template_txt_html.safe_substitute({"html_body": html_body})

    # HTMLファイル保存
    with open("./output/checklist.html", "w", encoding="utf-8") as f:
        f.write(template_txt_html)

    messagebox.showinfo('Info', "HTMLファイルが生成されました")

# カラム選択ボタンを表示
btn_select_columns = tk.Button(root, text="カラムを選択", command=show_column_selection_dialog)
btn_select_columns.pack()

root.mainloop()

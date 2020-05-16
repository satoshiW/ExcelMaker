import os, sys
import pandas as pd
import openpyxl as px
import tkinter as tk
import tkinter.filedialog as fl
import tkinter.messagebox as mb
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.styles import PatternFill

#ファイル選択関数
def select_file():
    global file_path
    filetype = [("all file", "*.csv")]
    iDir = os.path.abspath(r"C:\Users\めっき実験系列\Desktop\測温データ")
    file_path = fl.askopenfilename(initialdir=iDir, filetypes=filetype)
    file1.set(file_path)

#エクセル作成関数
def make_excel():
    excel_name = str(os.path.splitext(file_path)[0]) + "-1.xlsx"
    df = pd.read_csv(file_path, skiprows=57, usecols=[2], encoding="cp932")
    df.drop(df.tail(3).index, inplace=True)
    df_float = df.astype("float").round(1)

    #同名ファイルがある場合、上書き保存するか確認
    if os.path.isfile(excel_name):
        res = mb.askquestion("", "同名ファイルがあります。上書きしますか？")
        if res == "yes":
            df_float.to_excel(excel_name, header=False, index=False)
        elif res == "no":
            mb.showinfo("", "もう一度ファイル名を確認してください")
            sys.exit()
    else:
        df_float.to_excel(excel_name, header=False, index=False)

    wb = px.load_workbook(excel_name)
    ws = wb.active
    sheet = wb["Sheet1"]
    sc = sheet.cell
    wc = ws.cell

    sheet.insert_cols(0, 1) #先頭に1列挿入

    start = 1 #昇温開始時間
    cell_diff1 = 0

    #上下のセルを比べ、3回連続で3以上上昇した場合昇温開始とする
    while cell_diff1 <= 3:
        start += 1
        cell_diff3 = float(sc(row=start+1, column=2).value) - float(sc(row=start, column=2).value)
        if cell_diff3 >= 3:
            cell_diff2 = float(sc(row=start+2, column=2).value) - float(sc(row=start+1, column=2).value)
            if cell_diff2 >= 3:
                cell_diff1 = float(sc(row=start+3, column=2).value) - float(sc(row=start+2, column=2).value)

    end = start #データの最終行
    v1 = 0
    
    #昇温時間を0.5ずつ入力
    while sc(row=end, column=2).value is not None:
        wc(row=end, column=1, value=v1)
        end += 1
        v1 += 0.5

    keep = start #保定開始時間
    fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    temp_var = int(entry_temp.get()) - 10 #狙い温度-10℃
    
    #狙い温度-10℃の行
    while sc(row=keep, column=2).value <= temp_var:
        keep += 1

    #小数点第一位が5の場合、1行下げる
    if str(sc(row=keep, column=1).value)[-1] == str(5):
        keep = keep + 1
        
    #保定開始時間のセルに色を付ける
    wc(row=keep, column=1).fill = fill
    wc(row=keep, column=2).fill = fill

    v2 = 0
    
    #狙い温度-10℃から、保定時間を0.5ずつ入力
    while keep != end:
        wc(row=keep, column=3, value=v2)
        keep += 1
        v2 += 0.5
        #該当の保定時間に色を付ける
        if int(entry_time1.get()) == v2 or int(entry_time2.get()) == v2 or int(entry_time3.get()) == v2:
            wc(row=keep, column=1).fill = fill
            wc(row=keep, column=2).fill = fill
            wc(row=keep, column=3).fill = fill
            
            max_entry_time = keep

    #セルの書式を小数第一位で揃える
    for row in sheet:
        for cell in row:
            cell.number_format = "0.0"

    #チャートの作成
    chart = ScatterChart()
    
    x_values = Reference(ws, min_row=start, min_col=1, max_row=end, max_col=1) #x軸（昇温時間）
    y_values = Reference(ws, min_row=start, min_col=2, max_row=end, max_col=2) #y軸（温度）

    graph = Series(y_values, x_values)
    chart.series.append(graph)
    
    ws.add_chart(chart, "D"+str(max_entry_time)) #チャートを保定終了時の行に表示
    
    wb.save(excel_name) #Excelファイルを保存
    mb.showinfo("", "Excelファイルを作成しました")

#GUIの作成
if __name__ == "__main__":
    root = tk.Tk()
    root.title("CSVをExcelに変換")

    #frame1
    frame1 = tk.LabelFrame(root, text="ファイルを選択")
    frame1.grid(row=0, columnspan=2, sticky="we", padx=5)

    select_button = tk.Button(frame1, text="選択", command=select_file, width=10)
    select_button.grid(row=0, column=3)

    #ファイルパスの表示
    file1 = tk.StringVar()
    file1_entry = tk.Entry(frame1, textvariable=file1, width=35)
    file1_entry.grid(row=0, column=2, padx=5)

    #frame2
    frame2 = tk.LabelFrame(root, text="条件")
    frame2.grid(row=1, sticky="we")

    text_temp = tk.Label(frame2, text="狙い温度（℃）", width=20)
    text_temp.grid(row=0, column=0, padx=5)

    text_time = tk.Label(frame2, text="保定時間（秒）:複数指定可", width=25)
    text_time.grid(row=0, column=1)

    action_button = tk.Button(frame2, text="実行", command=make_excel, width=15)
    action_button.grid(row=3, column=0)

    entry_temp = tk.Entry(frame2, width=15)
    entry_temp.grid(row=1, column=0, padx=5)

    entry_time1 = tk.Entry(frame2, width=15)
    entry_time1.grid(row=1, column=1, padx=5, pady=5)
    entry_time1.insert(tk.END, 0)

    entry_time2 = tk.Entry(frame2, width=15)
    entry_time2.grid(row=2, column=1, padx=5, pady=5)
    entry_time2.insert(tk.END, 0)

    entry_time3 = tk.Entry(frame2, width=15)
    entry_time3.grid(row=3, column=1, padx=5, pady=5)
    entry_time3.insert(tk.END, 0)

root.mainloop()

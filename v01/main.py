import pandas as pd
import time
import takeNumberAndVersion
import Crawler_ASTM
import Crawler_ISO
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

link = ""

def getLink():
    global link 
    link = filedialog.askopenfilename(
        title="選擇 Excel 檔案",
        filetypes=[("Excel 檔案", "*.xlsx *.xls")]
    )
    excelName.set(link.split("/")[len(link.split("/"))-1])
    print("清單位址:", link)
    
    return link

def run(link, _type):
    timeStart = time.time()
    print("執行")
    # _type = btn_type["text"]
    print("選擇類型", _type)
    OriginTable = takeNumberAndVersion.run(link, _type)
    print(_type, "標準共有", len(OriginTable["Standard Number"]), "筆資料")
    for x in OriginTable["Standard Number"]: print(x)
    match(_type):
        case "ASTM":        
            Crawler_ASTM.run(OriginTable)
        case "ISO":
            Crawler_ISO.run(OriginTable)
    timeEnd = time.time()
    durationTime = timeEnd-timeStart
    print(f"執行時間: {durationTime//60}分{durationTime%60}秒")
    
window = tk.Tk()
window.title("法規標準更新追蹤")
frame_readExcel = tk.Frame(window).pack(pady=5, padx=20)
excelName = tk.StringVar()
excelName.set("尚未選擇檔案")
label_ExcelName = ttk.Label(frame_readExcel, textvariable=excelName).pack()
btn_readExcel = ttk.Button(frame_readExcel, text="選擇檔案", command=getLink)
btn_readExcel.pack()
frame_standardType = tk.Frame(window).pack(pady=5, padx=20)
label_chooseStandardType = ttk.Label(frame_standardType, text="選擇檢查標準類型:").pack(anchor="w", padx=10)

standardtypeList = ['ASTM', 'ISO', '3', '4','5','6','7','8','9','10','11','12', '13', '14']
num = len(standardtypeList)
frame_typeBtn = tk.Frame(frame_standardType)
frame_typeBtn.pack(padx=10)
row = 0
col = 0
for i in range(len(standardtypeList)):
    if row%2==0:
        btn_type = ttk.Button(frame_typeBtn, text=standardtypeList[i], command= lambda t=standardtypeList[i]: run(link, t))
        btn_type.grid(row=row, column=col)
        col +=1
        while col==4: 
            col=0
            row+=1
        continue
    if row%2==1:
        btn_type = ttk.Button(frame_typeBtn, text=standardtypeList[i], command= lambda t=standardtypeList[i]: run(link, t))
        btn_type.grid(row=row, column=col, columnspan=2, pady=2, padx=2)
        col +=1
        while col==3: 
            col=0
            row+=1

window.mainloop()
#D:\UnaKuo\RA Project\Standards_Review\v01\StandardsList.xlsx
# link = input("輸入清單位址: ")
# print("清單位址", link)
# _type = input("請輸入檢查標準類型: ")





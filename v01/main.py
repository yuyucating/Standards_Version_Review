import pandas as pd
import takeNumberAndVersion
import Crawler_ASTM
import Crawler_ISO
   

#D:\UnaKuo\RA Project\Standards_Review\v01\StandardsList.xlsx
link = input("輸入清單位址: ")
print("清單位址", link)
_type = input("請輸入檢查標準類型: ")

OriginTable = takeNumberAndVersion.run(link, _type)
print(_type, "標準共有", len(OriginTable["Standard Number"]), "筆資料")
for x in OriginTable["Standard Number"]: print(x)


match(_type):
    case "ASTM":        
        Crawler_ASTM.run(OriginTable)
    case "ISO":
        Crawler_ISO.run(OriginTable)
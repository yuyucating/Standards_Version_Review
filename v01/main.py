import takeNumberAndVersion
import Crawler_ASTM

#D:\UnaKuo\RA Project\Standards_Review\v01\StandardsList.xlsx
link = input("輸入清單位址: ")
print("清單位址", link)
type = input("請輸入檢查標準類型: ")



match(type):
    case "ASTM":
        OriginTable_ASTM = takeNumberAndVersion.run(link, type)
        print("ASTM 標準共有", len(OriginTable_ASTM), "筆資料")
        
        Crawler_ASTM.run(OriginTable_ASTM)
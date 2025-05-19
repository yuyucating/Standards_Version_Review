import takeNumberAndVersion

#D:\UnaKuo\RA Project\Standards_Review\v01\StandardsList.xlsx
link = input("輸入清單位址: ")
print("清單位址", link)
type = input("請輸入檢查標準類型: ")

takeNumberAndVersion.run(link, type)
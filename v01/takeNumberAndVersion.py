import pandas as pd

def splitStandardNumberAndVersion(standard, divider):
    print(standard, "切分資料量: ", len(standard.split(divider)))
    standardNumber = standard.split(divider)[0]
    standardVersion = standard.split(divider)[1]
    return(standardNumber, standardVersion)

def getStandardNumbersAndVersion(documentNumbers, originalVersions, _type):
    print("正在從 excel 取得標準編號及收錄版本...")
    print("共",len(documentNumbers),"筆")
    print(documentNumbers)
    standardNumbers = []
    standardVersions = []
    
    match(_type):
        case "ISO":
            i=0
            for standard in documentNumbers:
                if ":" in standard:
                    # 包含年份/版本
                    standardNumber, standardVersion = splitStandardNumberAndVersion(standard, ":")
                elif "：" in standard:
                    # 包含年份/版本
                    standardNumber, standardVersion = splitStandardNumberAndVersion(standard, "：")
                else: 
                    # 不包含年份/版本
                    standardNumber = standard
                    standardVersion = originalVersions[i]
                standardNumbers.append(str(standardNumber))
                standardVersions.append(str(standardVersion))
                i+=1
                print("完成第",i,"筆 / 共",len(documentNumbers),"筆")
                print(">> 取得", standardNumber, standardVersion)
            print(_type,"擷取完成!")
            
        case "ASTM":
            i=0
            for standard in documentNumbers:
                try: 
                    if "-" in standard:
                        # 包含年份/版本
                        # 判斷編碼登記格式
                        if standard.startswith("ASTM-"):
                            print("[[ASTM修正]]")
                            standard = standard.replace("ASTM-", "ASTM ")
                        elif standard.startswith("ASTM_"):
                            print("[[ASTM修正]]")
                            standard = standard.replace("ASTM_", "ASTM ")
                        if standard.count("-")!=1:
                            n = standard.find("-") # 取得第1個 "-" 的位置
                            while standard[n+1].isalpha():
                                standard = standard.replace("-", "/", 1)
                                n = standard.find("-")
                            
                    standardNumber, standardVersion = splitStandardNumberAndVersion(standard, "-")
                
                    #判斷版本登記格式
                    if " R" in standardVersion:
                        if "e" in standardVersion:
                            standardVersion = standardVersion.replace(" R", "(").replace("e", ")e")
                        else: standardVersion = standardVersion.replace(" R", "(")+")"      
                except Exception as e:
                    standardNumber = standard
                    print("Standard Number is", standardNumber)
                    
                standardNumbers.append(str(standardNumber))
                standardVersions.append(str(standardVersion))
                    
                i+=1
                print("完成第",i,"筆 / 共",len(documentNumbers),"筆")
                print(">> 取得", standardNumber, standardVersion)
            print(_type,"擷取完成!")
                    
    return standardNumbers, standardVersions      
            
    

def run(link, _type):
    if ".xlsx" in link:
        sheetData = pd.read_excel(link, sheet_name=_type, engine="openpyxl")
    elif ".xls" in link:
        sheetData = pd.read_excel(link, sheet_name=_type, engine="xlrd")
    else: print("檔案格式錯誤")
        
    for col in sheetData:
        if "Document Number" in col:
            originalDocumentNumbers = col
        if "Version" in col:
            originalStandardVersions = col

    standardNumbers, standardVersions = getStandardNumbersAndVersion(sheetData[originalDocumentNumbers], sheetData[originalStandardVersions], _type)

    print("回傳兩個檔資料格式分別為", type(standardNumbers), type(standardVersions))
    newTable = pd.DataFrame()
    newTable["Standard Number"] = standardNumbers
    newTable["Registed Version"] = standardVersions
    print(newTable)
    print("=== 離開 takdNumberAndVersion.py ===")
    return newTable
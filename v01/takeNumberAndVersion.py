import pandas as pd

def splitStandardNumberAndVersion(standard, divider):
    # print("切分資料量: ", len(standard.rsplit(divider)))
    standardNumber = standard.rsplit(divider,1)[0]
    standardVersion = standard.rsplit(divider,1)[1]
    return(standardNumber, standardVersion)

def getStandardNumbersAndVersion(documentNumbers, originalVersions, _type):
    print("正在從 excel 取得標準編號及收錄版本...")
    print("共",len(documentNumbers),"筆")
    print(documentNumbers)
    standardTypes = []
    standardNumbers = []
    standardVersions = []
    
    match(_type):
        case "ISO":
            i=0
            for standard in documentNumbers:
                # 格式化標準編號與版本
                if "amd" in str(standard).lower():
                    if "：" in standard:
                        standard = standard.replace("：", ":")
                    splitMark = standard.upper().find("AMD")
                    standard = standard[:splitMark-1]
                    standardTypes.append("Amd")
                elif "cor" in str(standard).lower():
                    if "：" in standard:
                        standard = standard.replace("：", ":")
                    splitMark = standard.upper().find("COR")
                    standard = standard[:splitMark-1]
                    standardTypes.append("Cor")
                else:
                    if "：" in str(standard):
                        standard = standard.replace("：", ":")
                        standard = standard.split(":")[0]
                    print(standard, ">> main article")
                    standardTypes.append("Main")
                try:
                    standardVersion = int(originalVersions[i])
                except Exception as e:
                    print("nan PASS!!!")              
                standardNumbers.append(str(standard).strip())
                standardVersions.append(str(standardVersion).strip())
                i+=1
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
                    
                standardTypes.append("Main")
                standardNumbers.append(str(standardNumber))
                standardVersions.append(str(standardVersion))
                    
                i+=1
                print(">> 取得", standardNumber, standardVersion)
            print(_type,"擷取完成!")
        case "YY":            
            for standard in documentNumbers:
                try:
                    # 辨別標準類型是否包含空格(Typo)
                    if len(standard.split(" ")) ==3:
                        standard = standard.split(" ")[0]+standard.split(" ")[1]+" "+standard.split(" ")[2]
                    if isinstance(standard, str) and "／" in standard:
                        standard = standard.replace("／", "/")
                        standardTypes.append("YY/T")
                    elif isinstance(standard, str) and "/" in standard:
                        standardTypes.append("YY/T")
                    else: standardTypes.append("YY")
                    standardNumbers.append(standard)
                except:
                    pass
            standardVersions = originalVersions
                
    return standardTypes, standardNumbers, standardVersions      
            
    

def run(link, _type):
    print("=== 執行 takdNumberAndVersion.py ===")
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

    standardTypes, standardNumbers, standardVersions = getStandardNumbersAndVersion(sheetData[originalDocumentNumbers], sheetData[originalStandardVersions], _type)

    newTable = pd.DataFrame()
    newTable["Type"] = standardTypes
    newTable["Standard Number"] = standardNumbers
    newTable["Registed Version"] = standardVersions
    print(newTable)
    print("=== 離開 takdNumberAndVersion.py ===")
    return newTable
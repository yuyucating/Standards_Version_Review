import pandas as pd

def splitStandardNumberAndVersion(standard, divider):
    print(standard, "切分資料量: ", len(standard.split(divider)))
    standardNumber = standard.split(divider)[0]
    standardVersion = standard.split(divider)[1]
    return(standardNumber, standardVersion)

def run(link, _type):
    inputStandardNumbers = pd.read_excel(link, sheet_name=_type, usecols=["Document Number"])
    print("從 excel 取得標準編號:")
    print(inputStandardNumbers)
    print(f"一共有{len(inputStandardNumbers)}筆資料")

    print("分析標準編號與版本")
    newStandardNumbers = inputStandardNumbers
    splitedNumberList = []
    splitedVersionList = []

    match(_type):
        case "ISO":
            for standard in newStandardNumbers["Document Number"]:    
                print()
                if(":" in standard):
                    standardNumber, standardVersion = splitStandardNumberAndVersion(standard, ":")
                    newStandardNumbers.loc[newStandardNumbers["Document Number"]==standard, "Standard Number"] = standardNumber
                    newStandardNumbers.loc[newStandardNumbers["Document Number"]==standard, "Standard Version"] = standardVersion
                elif("：" in standard):
                    standardNumber, standardVersion = splitStandardNumberAndVersion(standard, "：")
                    newStandardNumbers.loc[newStandardNumbers["Document Number"]==standard, "Standard Number"] = standardNumber
                    newStandardNumbers.loc[newStandardNumbers["Document Number"]==standard, "Standard Version"] = standardVersion

        case "ASTM":
            for standard in newStandardNumbers["Document Number"]:
                print(standard)
                # 判斷編碼登記格式
                if standard.startswith("ASTM-"):
                    print("[[ASTM修正]]")
                    standard = standard.replace("ASTM-", "ASTM ")
                elif standard.startswith("ASTM_"):
                    print("[[ASTM修正]]")
                    standard = standard.replace("ASTM_", "ASTM ")
                    
                if standard.count("-")!=1:
                    i = standard.find("-") # 取得第1個 "-" 的位置
                    while standard[i+1].isalpha():
                        standard = standard.replace("-", "/", 1)
                        i = standard.find("-")
                
                standardNumber, standardVersion = splitStandardNumberAndVersion(standard, "-")
                print("分割完成!! 取得", standardNumber, standardVersion)
                
                # 判斷版本登記格式
                if " R" in standardVersion:
                    if "e" in standardVersion:
                        standardVersion = standardVersion.replace(" R", "(").replace("e", ")e")
                    else: standardVersion = standardVersion.replace(" R", "(")+")"
                        
                splitedNumberList.append(standardNumber)
                splitedVersionList.append(standardVersion)

    newStandardNumbers["Standard Number"] = splitedNumberList
    newStandardNumbers["Standard Version"] = splitedVersionList
    print("現在取得的資料為: ", _type)
    print(newStandardNumbers)
    
    return newStandardNumbers
import pandas as pd

def splitStandardNumberAndVersion(standard, divider):
    print(standard, "切分資料量: ", len(standard.split(divider)))
    standardNumber = standard.split(divider)[0]
    standardVersion = standard.split(divider)[1]
    return(standardNumber, standardVersion)

def run(link, type):
    inputStandardNumbers = pd.read_excel(link, sheet_name=type, usecols=["Document Number"])
    print("從 excel 取得標準編號:")
    print(inputStandardNumbers)
    print(f"一共有{len(inputStandardNumbers)}筆資料")

    print("分析標準編號與版本")
    newStandardNumbers = inputStandardNumbers

    match(type):
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
                standardNumber, standardVersion = splitStandardNumberAndVersion(standard, "-")
                newStandardNumbers.loc[newStandardNumbers["Document Number"]==standard, "Standard Number"] = standardNumber
                newStandardNumbers.loc[newStandardNumbers["Document Number"]==standard, "Standard Version"] = standardVersion

    print("現在取得的資料為: ", type)
    print(newStandardNumbers)
import selenium
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import date

def is_excel_file_open(file_path):
    try:
        # 嘗試以唯讀寫入模式打開
        with open(file_path, "r+b"):
            return False  # 沒有被佔用，代表沒開啟
    except IOError:
        return True  # 被佔用，可能是開啟中

def run(originalTable):
    print("執行 Crawler_ASTM.py")
    currentVersionList = []
    isUpdateList = []
    
    driver = webdriver.Chrome()

    i=0
    print("資料長度", len(originalTable["Standard Number"]))
    
    for standard in originalTable["Standard Number"]:
        link_web = "https://www.astm.org/search/result?q="+standard.replace(" ","%20")
        driver.get(link_web)
        print('。開始搜尋', standard, "(第", i+1, "項/共", len(originalTable["Standard Number"]), "項)")
     
        try:
            onlyNumber = standard.split()[1]
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//p[contains(text(), {onlyNumber})]")))
            print(f"找到 {onlyNumber} !!")
            onlyNumber_temp = onlyNumber+"-"
            SearchedStandard = driver.find_element(By.XPATH, f"//*[contains(@class, 'searchComponent_sku__V2OCP') and starts-with(text(), '{onlyNumber_temp}')]").get_attribute("textContent")
            SearchedVersion = SearchedStandard.split("-")[1]
            print("現行版本是:", SearchedVersion)
            
            currentVersionList.append(SearchedVersion)
                        
            if SearchedVersion.strip()==originalTable.iloc[i]["Registed Version"].strip():
                print("沒有改版")
                isUpdateList.append("-")
            else:
                print("改版!")
                isUpdateList.append("!! UPDATED !!")
        except Exception as e:
            print("第一個 try 出錯了：", e)
            try:
                print(f"找不到 {onlyNumber} 換一個關鍵字進行搜尋...")
                link_web = "https://www.astm.org/search/result?q="+standard.replace(" ","%20")+"-"+originalTable["Registed Version"][i]
                driver.get(link_web)
                print('搜尋', standard+"-"+originalTable["Registed Version"][i])
                onlyNumber = standard.split()[1]
                SearchedNumber = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//p[contains(text(), {onlyNumber})]")))
                
                print(f"找到 {onlyNumber}"+"-"+originalTable["Registed Version"][i], "!!")
                onlyNumber_temp = onlyNumber+"-"
                SearchedStandard = driver.find_element(By.XPATH, f"//*[contains(@class, 'searchComponent_sku__V2OCP') and starts-with(text(), '{onlyNumber_temp}')]").get_attribute("textContent")
                SearchedVersion = SearchedStandard.split("-")[1]
                print("現行版本是:", SearchedVersion, "第", i+1, "項/共", len(originalTable["Standard Number"]), "項")
                
                currentVersionList.append(SearchedVersion)
                        
                if SearchedVersion.replace(" ","")==originalTable.iloc[i]["Registed Version"].replace(" ",""):
                    print("沒有改版")
                    isUpdateList.append("-")
                else:
                    print("改版!")
                    isUpdateList.append("!! UPDATED !!")
                
            except Exception as e:
                print("第二個 try 出錯了：", e)
                print("...沒有執行成功")
                currentVersionList.append("未成功執行")
                isUpdateList.append("未成功執行")
            
        i+=1
            
    driver.quit()
    newTable = originalTable
    newTable["Current Version"] = currentVersionList
    newTable["Update"] = isUpdateList
    
    print(newTable)  
      
    print("準備輸出 excel")
    while is_excel_file_open(f"檔案_{date.today()}.xlsx"):
        input("輸出 excel 開啟中，無法存檔!!關閉檔案後，隨意輸入，再次嘗試存檔。")
    newTable.to_excel(f"檔案_{date.today()}.xlsx", sheet_name="ASTM", index=False) 
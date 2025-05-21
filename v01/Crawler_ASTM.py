import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import date

def run(originalTable):
    print("執行 Crawler_ASTM.py")
    currentVersionList = []
    isUpdateList = []
    
    driver = webdriver.Chrome()

    i=0
    for standard in originalTable["Standard Number"]:
        
        link_web = "https://www.astm.org/search/result?q="+standard.replace(" ","%20")
        driver.get(link_web)
        print('搜尋', standard)
     
        try:
            onlyNumber = standard.split()[1]
            SearchedNumber = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//p[contains(text(), {onlyNumber})]")))
            print(f"找到 {onlyNumber} !!")
            onlyNumber_temp = onlyNumber+"-"
            SearchedStandard = driver.find_element(By.XPATH, f"//*[contains(@class, 'searchComponent_sku__V2OCP') and starts-with(text(), '{onlyNumber_temp}')]").get_attribute("textContent")
            SearchedVersion = SearchedStandard.split("-")[1]
            print("現行版本是:", SearchedVersion, "第", i+1, "項/共", len(originalTable["Standard Number"]), "項")
            
            currentVersionList.append(SearchedVersion)
                       
            if SearchedVersion.strip()==originalTable.iloc[i]["Standard Version"].strip():
                print("沒有改版")
                isUpdateList.append("-")
            else:
                print("改版!")
                isUpdateList.append("!! UPDATED !!")
        except:
            try:
                print("找不到... 換一個關鍵字進行搜尋...")
                link_web = "https://www.astm.org/search/result?q="+standard.replace(" ","%20")+"-"+originalTable["Standard Version"][i]
                driver.get(link_web)
                print('搜尋', standard+"-"+originalTable["Standard Version"][i])
                onlyNumber = standard.split()[1]
                SearchedNumber = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//p[contains(text(), {onlyNumber})]")))
                
                print(f"找到 {onlyNumber}"+"-"+originalTable["Standard Version"][i], "!!")
                onlyNumber_temp = onlyNumber+"-"
                SearchedStandard = driver.find_element(By.XPATH, f"//*[contains(@class, 'searchComponent_sku__V2OCP') and starts-with(text(), '{onlyNumber_temp}')]").get_attribute("textContent")
                SearchedVersion = SearchedStandard.split("-")[1]
                print("現行版本是:", SearchedVersion, "第", i+1, "項/共", len(originalTable["Standard Number"]), "項")
                
                currentVersionList.append(SearchedVersion)
                        
                if SearchedVersion.strip()==originalTable.iloc[i]["Standard Version"].strip():
                    print("沒有改版")
                    isUpdateList.append("-")
                else:
                    print("改版!")
                    isUpdateList.append("!! UPDATED !!")
                
            except:
                print("...沒有執行成功")
                currentVersionList.append("未成功執行")
                isUpdateList.append("未成功執行")
            
        i+=1
            
    driver.quit()
    newTable = originalTable
    newTable["Current Version"] = currentVersionList
    newTable["Update"] = isUpdateList
    
    print(newTable)  
    newTable.to_excel(f"檔案_{date.today()}.xlsx", sheet_name="ASTM", index=False)   
    print("已輸出 excel")
    
import selenium
import os
import pandas as pd
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

def run(originTable):
    driver = webdriver.Chrome()

    standardNumbers = originTable["Standard Number"]
    currentVersionList = []
    isUpdateList = []

    link_web = "https://www.iso.org/home.html"
    driver.get(link_web)
    print("進入 ISO 官網")
    
    try:
        cookie_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))
        )
        cookie_button.click()
    except:
        print("No cookie popup appeared.")

    i=0
    for standardNumber in standardNumbers:
        SearchingBox = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "form-control")))
        submitBtn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".bi-search.text-muted")))
        
        SearchingBox.clear()
        SearchingBox.send_keys(standardNumber)
        submitBtn.click()
        print("搜尋", standardNumber)

        WebDriverWait(driver, 10).until(
            EC.text_to_be_present_in_element((By.ID, "stats-container"), "results found")
        )
        result = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@id='stats-container']")))
        print(result.text)
        if "0 results" in result.text:
            print(f"沒有找到 {standardNumber}")
            currentVersionList.append("可能廢止")
            isUpdateList.append("-")
        else: 
            searchedStandard = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//a[starts-with(normalize-space(.), '{standardNumber}')]")))

            print(f"找到 {standardNumber} !!")
            currentVersion = searchedStandard.get_attribute("textContent").split(":")[1]
            print(f"現行版本是: {currentVersion}")
            
            if originTable["Registed Version"][i]==currentVersion:
                isUpdateList.append("-")
            else:
                isUpdateList.append("!! UPDATED !!")
            currentVersionList.append(currentVersion)
    
    newTable = originTable
    newTable["Current Version"] = currentVersionList
    newTable["Update"] = isUpdateList

    print(originTable)

    driver.quit()
    
    print("準備輸出 excel")
    excel_path = f"法規標準更新檢查_{date.today()}.xlsx"
    if os.path.exists(excel_path):
        print("檔案存在, 準備寫入...")
        while is_excel_file_open(excel_path):
            input("輸出 excel 開啟中，無法存檔!!關閉檔案後，隨意輸入，再次嘗試存檔。")
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="new") as writer: newTable.to_excel(writer, sheet_name="ISO", index=False)
    else:
        print(f"建立: 檔案_{date.today()}.xlsx")
        newTable.to_excel(excel_path, sheet_name="ISO", index=False)
    
    
    print("已輸出 excel")
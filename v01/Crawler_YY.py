import selenium
from selenium import webdriver
import os
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from datetime import datetime
from datetime import date

def is_excel_file_open(file_path):
    try:
        # 嘗試以唯讀寫入模式打開
        with open(file_path, "r+b"):
            return False  # 沒有被佔用，代表沒開啟
    except IOError:
        return True  # 被佔用，可能是開啟中
    
def crawler(driver, SearchingBox, submitBtn, standardList, registedVersionList, versionList, updateList, standard, registedVersion):
    SearchingBox.clear()
    SearchingBox.send_keys(standard)
    submitBtn.click()
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "hbtable"))
    )
    print("開始搜尋: "+standard)

    # 找到 <tbody> 元素
    # tbody = WebDriverWait(driver, 10).until(
    #     EC.presence_of_element_located((By.TAG_NAME, "tbody")))
    tbody = driver.find_element(By.ID, "hbtable").find_element(By.TAG_NAME, "tbody")

    # 找出所有列（tr）
    rows = tbody.find_elements(By.TAG_NAME, "tr")
    print(standard, "有", len(rows), "行")

    for i in range(len(rows)):
        try:
            # 每次重新定位 row，避免舊元素失效
            fresh_row = WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, "tr")))[i]
            cells = fresh_row.find_elements(By.TAG_NAME, "td")
            data = []
            for cell in cells:
                data.append(cell.text.strip())
            print("。。", data)
            _temp = data[1].replace("—", "-").replace("–", "-").replace("\xa0", " ").strip()
            currentVersion = _temp[-4:]
            establishDate = data[6]
            d = establishDate.split("-")
            print("Registed Version:", registedVersion, "\nCurrent Version:", currentVersion)
            if str(currentVersion) != registedVersion:
                today = date.today()
                if today >= date(int(d[0]), int(d[1]), int(d[2])):
                    print("!!UPDATE!!")
                    updateList.append("!!UPDATE!!")
                else: 
                    print("Update in nearly future")
                    updateList.append("Update in nearly future.")
            else:
                updateList.append("-")
            standardList.append(standard)
            registedVersionList.append(registedVersion)
            versionList.append(currentVersion)

            
        except Exception as e:
            pass
    
    print(len(standardList), len(registedVersionList), len(versionList), len(updateList))
    return pd.DataFrame({"Standard Number":standardList, "Registed Version": registedVersionList, "Current Version": versionList, "Update": updateList})
    

def run(originTable):
    standardList = []
    registedVersionList = []
    versionList = []
    updateList = []

    print("=== 執行 Crawler_YY.py ===")
    driver = webdriver.Chrome()

    link_web = "https://hbba.sacinfo.org.cn/stdList"
    driver.get(link_web)
    print("進入中國行業標準信息服務平台-標準查詢")

    SearchingBox = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "keyword")))
    submitBtn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "searchBtn")))

    standard = originTable["Standard Number"]
    registedVersion = originTable["Registed Version"]
    
    for i in range(len(originTable)):
        newTable = crawler(driver, SearchingBox, submitBtn, standardList, registedVersionList, versionList, updateList, standard[i], registedVersion[i])
    driver.quit()
    
    print(newTable)

    print("準備輸出 excel")
    excel_path = f"法規標準更新檢查_{date.today()}.xlsx"
    if os.path.exists(excel_path):
        print("檔案存在, 準備寫入...")
        while is_excel_file_open(excel_path):
            input("輸出 excel 開啟中，無法存檔!!關閉檔案後，隨意輸入，再次嘗試存檔。")
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="new") as writer: newTable.to_excel(writer, sheet_name="YY", index=False)
    else:
        print(f"建立: 檔案_{date.today()}.xlsx")
        newTable.to_excel(excel_path, sheet_name="YY", index=False)


    print("已輸出 excel")
    print("=== 離開 Crawler_YY.py ===")
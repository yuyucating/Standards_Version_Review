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
import traceback

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
    print("開始搜尋: "+ str(standard))

    try:
            # 找到 <tbody> 元素
        WebDriverWait(driver, 10).until(
            EC.text_to_be_present_in_element((By.ID, "hbtable"), standard)
            )

        tbody =  driver.find_element(By.ID, "hbtable").find_element(By.TAG_NAME, "tbody")
        
        # 找出所有列（tr）
        rows = tbody.find_elements(By.TAG_NAME, "tr")
        print(standard, "有", len(rows), "行")

    
        for i, row in enumerate(rows):
            # try:
            cells = row.find_elements(By.TAG_NAME, "td")
            data = [cell.text.strip() for cell in cells]
            
            if len(data) < 7:
                print(f"資料不足：{data}")
                continue
            
            print("。。", data)
            _temp = data[1].replace("—", "-").replace("–", "-").replace("\xa0", " ").strip()
            print("抓抓號碼", _temp[:-5],"\n註冊號碼", standard)
            if str(_temp[:-5]) != str(standard):
                print("不是", standard,", 跳過!!!")
                continue
            currentVersion = _temp[-4:]
            establishDate = data[6]
            d = establishDate.split("-")
            print("Registed Version:", registedVersion, "\nCurrent Version:", currentVersion)
            
            if str(currentVersion) != str(registedVersion):
                today = date.today()
                if today >= date(int(d[0]), int(d[1]), int(d[2])):
                    print("!!UPDATE!!")
                    updateList.append("!!UPDATE!!")
                else:
                    print("Update in nearly future")
                    updateList.append("Update in nearly future.")
                standardList.append(standard)
                registedVersionList.append(registedVersion)
                versionList.append(currentVersion)
            else:
                updateList.append("-")
                standardList.append(standard)
                registedVersionList.append(registedVersion)
                versionList.append(currentVersion)
                break
    except Exception as e:
        print("查不到 換一個標準類型進行搜尋...")
        splited = standard.split(" ")
        if "T" in splited:
            _standard = "YY "+splited[1]
        else:
            _standard = "YY/T "+splited[1]
        SearchingBox.clear()
        SearchingBox.send_keys(_standard)
        submitBtn.click()
        print("開始搜尋: "+ str(_standard))
        try:
            WebDriverWait(driver, 10).until(
                EC.text_to_be_present_in_element((By.ID, "hbtable"), _standard)
                )

            tbody =  driver.find_element(By.ID, "hbtable").find_element(By.TAG_NAME, "tbody")
            
            # 找出所有列（tr）
            rows = tbody.find_elements(By.TAG_NAME, "tr")
            print(_standard, "有", len(rows), "行")

        
            for i, row in enumerate(rows):
                # try:
                cells = row.find_elements(By.TAG_NAME, "td")
                data = [cell.text.strip() for cell in cells]
                
                if len(data) < 7:
                    print(f"資料不足：{data}")
                    continue
                
                print("。。", data)
                _temp = data[1].replace("—", "-").replace("–", "-").replace("\xa0", " ").strip()
                print("抓抓號碼", _temp[:-5],"\n註冊號碼", standard)
                if _temp[:-5] != standard:
                    print("不是", standard,", 跳過!!!")
                currentVersion = _temp[-4:]
                establishDate = data[6]
                d = establishDate.split("-")
                print("Registed Version:", registedVersion, "\nCurrent Version:", currentVersion)
                
                if str(currentVersion) != str(registedVersion):
                    today = date.today()
                    if today >= date(int(d[0]), int(d[1]), int(d[2])):
                        print("!!UPDATE!! 且更新為", _standard)
                        updateList.append("!!UPDATE!! 且更新為 "+_standard)
                    else:
                        print(_standard, "update in nearly future")
                        updateList.append(_standard+" update in nearly future.")
                    standardList.append(standard)
                    registedVersionList.append(registedVersion)
                    versionList.append(currentVersion)
                else:
                    updateList.append("更換標準類型, 但未進行改版: "+_standard)
                    standardList.append(standard)
                    registedVersionList.append(registedVersion)
                    versionList.append(currentVersion)
                    break
        except Exception as e:
            standardList.append(standard)
            registedVersionList.append(registedVersion)
            versionList.append("查無資料")
            updateList.append("請自行確認是否廢止")
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
        print(f"完成第{i}個/總共{len(originTable)+1}個")
    driver.quit()
    
    print(newTable)

    print("準備輸出 excel")
    excel_path = f"法規標準更新檢查_{date.today()}.xlsx"
    if os.path.exists(excel_path):
        print("檔案存在, 準備寫入...")
        while is_excel_file_open(excel_path):
            input("輸出 excel 開啟中，無法存檔!!關閉檔案後按 Enter 嘗試存檔。")
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="new") as writer: newTable.to_excel(writer, sheet_name="YY", index=False)
    else:
        print(f"建立: 檔案_{date.today()}.xlsx")
        newTable.to_excel(excel_path, sheet_name="YY", index=False)


    print("已輸出 excel")
    print("=== 離開 Crawler_YY.py ===")
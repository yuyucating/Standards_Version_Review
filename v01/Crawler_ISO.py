import selenium
import findDownloadFolder
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import date

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

        searchedStandard = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//a[starts-with(normalize-space(.), '{standardNumber}')]")))

        print(f"找到 {standardNumber} !!")
        print(searchedStandard.get_attribute("textContent").split(":")[1])
        currentVersion = searchedStandard.get_attribute("textContent").split(":")[1]
        print(f"現行版本是: {currentVersion}")
        
        if originTable["Registed Version"][i]==currentVersion:
            isUpdateList = "-"
        else: isUpdateList = "!! UPDATED !!"
            
        currentVersionList.append(currentVersion)
    
    newTable = originTable
    newTable["Current Version"] = currentVersionList
    newTable["Update"] = isUpdateList

    print(originTable)

    driver.quit()
    
    newTable.to_excel(f"檔案_{date.today()}.xlsx", sheet_name="ISO", index=False)   
    print("已輸出 excel")
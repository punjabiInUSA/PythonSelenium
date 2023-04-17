import time
import openpyxl
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

# Create a new instance of Edge webdriver
mDriver = webdriver.Edge()
mWait = WebDriverWait(mDriver, 20)

# Add file path
# data_only marked as true so only values are fetched even if those are generated from formulas
excelWorkbook = openpyxl.load_workbook("PATH_TO_EXCEL_FILE_WITH_FILENAME_AND_EXTENSION", data_only=True)
# Excel Format expected as
# Country | Exit Date | Exit Month | Exit Year | Entry Date | Entry Month | Entry Year | Purpose  | Exit Full Date | Entry Full Date - Header Row
# Brazil  | 15        | April      | 2023      | 24         | May         | 2023       | Vacation | 4/15/2023      | 5/24/2023 - Data Row

# Add sheet name
targetSheetName = excelWorkbook.active

# Stores excel data
data = []

# Remove max_row=XX, if pulling complete data from excel
for row in targetSheetName.iter_rows(min_row=2, max_row=3, values_only=True):
    item = {
        'country': row[0],
        'exitDate': row[1],
        'exitMonth': row[2],
        'exitYear': row[3],
        'entryDate': row[4],
        'entryMonth': row[5],
        'entryYear': row[6],
        'purpose': row[7],
        # extracts just the date from original output 2023-04-15 00:00:00
        'exitFullDate': str(row[8]).split()[0],
        'entryFullDate': str(row[9]).split()[0]
    }

    # Prints are for debugging only
    # print(item['country'])
    # print(item['exitFullDate'])
    # print(item['entryFullDate'])
    data.append(item)

webAddress = "https://eservices.cic.gc.ca/rescalc/resCalcStartNew.do"
username = "ENTER_USER_ID"
password = "ENTER_PASSWORD"

# Navigate to the website
mDriver.get(webAddress)

# click retrieve saved calc button
retrieveSavedCalc = mDriver.find_element(By.XPATH, "//*[@id=\"wb-main-in\"]/div[1]/a[5]")
retrieveSavedCalc.click()

userField = mDriver.find_element(By.NAME, "username")
pwdField = mDriver.find_element(By.NAME, "password1")
btnLogin = mDriver.find_element(By.XPATH, "/html/body/main/form/section/div/div[6]/input")
userField.send_keys(username)
pwdField.send_keys(password)
btnLogin.click()
time.sleep(2)

paragraph = mWait.until(EC.
                        visibility_of_element_located((By.XPATH, "//*[@id=\"wb-main-in\"]/div[2]/div[1]/p/strong")))

if paragraph.is_displayed():
    # click retrieve saved calc button as website displays an error first time, requires re-initialization
    retrieveSavedCalc = mDriver.find_element(By.XPATH, "//*[@id=\"wb-main-in\"]/div[1]/a[5]")
    retrieveSavedCalc.click()

viewExistingCalc = mWait.until(
    EC.visibility_of_element_located((By.XPATH, "//*[@id=\"wb-main-in\"]/table/tbody/tr/td[3]/a[1]")))
viewExistingCalc.click()
time.sleep(2)
# btnModify = mDriver.find_element(By.XPATH, "//*[@id=\"wb-main-in\"]/form[1]/div[1]/input[4]")
btnModify = mDriver.find_element(By.NAME, "modify")
btnModify.click()
time.sleep(2)
# btnCalculate = mDriver.find_element(By.XPATH, "//*[@id=\"wb-main-in\"]/form/div/input[3]")
btnCalculate = mDriver.find_element(By.NAME, "next")
btnCalculate.click()
time.sleep(2)
# btnTempContinue = mDriver.find_element(By.XPATH, "//*[@id=\"wb-main-in\"]/form/div[1]/input[4]")
btnTempContinue = mDriver.find_element(By.NAME, "temprescontinue")
btnTempContinue.click()
time.sleep(2)
# btnPrisonContinue = mDriver.find_element(By.XPATH, "//*[@id=\"wb-main-in\"]/form/div[1]/input[4]")
btnPrisonContinue = mDriver.find_element(By.NAME, "prisoncontinue")
btnPrisonContinue.click()
time.sleep(2)

try:
    # will only process one row per argument data[:1], remove this to process all rows.
    for item in data:
        # Page Element Definitions
        # viewAbsenceDestination = mDriver.find_element(By.XPATH, "//*[@id=\"absenDestination\"]")
        viewAbsenceDestination = mDriver.find_element(By.ID, "absenDestination")
        # viewFromDate = mDriver.find_element(By.XPATH, "//*[@id=\"absenceFromDate\"]")
        viewFromDate = mDriver.find_element(By.ID, "absenceFromDate")
        # viewToDate = mDriver.find_element(By.XPATH, "//*[@id=\"absenceToDate\"]")
        viewToDate = mDriver.find_element(By.ID, "absenceToDate")
        # viewAbsenceReason =  mDriver.find_element(By.XPATH, "//*[@id=\"absencesReason\"]")
        viewAbsenceReason = mDriver.find_element(By.ID, "absencesReason")
        btnAddAbsence = mDriver.find_element(By.XPATH, "//*[@id=\"wb-main-in\"]/form/section[2]/div/div[4]/input[1]")
        btnSaveRecord = mDriver.find_element(By.NAME, "saveapp")

        # Focus onto absence entry area
        mDriver.execute_script("arguments[0].scrollIntoView();", viewAbsenceReason)

        # Select destination
        selectAbsenDest = Select(viewAbsenceDestination)
        selectAbsenDest.select_by_visible_text(item['country'])

        # Enter Date you left
        viewFromDate.clear()
        # viewFromDate.send_keys("04-15-2023") //debug
        exitFullDate = datetime.strptime(item['exitFullDate'], '%Y-%m-%d').strftime('%m-%d-%Y')
        item['exitFullDate'] = exitFullDate
        viewFromDate.send_keys(item['exitFullDate'])

        # Enter Date you returned
        viewToDate.clear()
        # viewToDate.send_keys("04-15-2023") //debug
        entryFullDate = datetime.strptime(item['entryFullDate'], '%Y-%m-%d').strftime('%m-%d-%Y')
        item['entryFullDate'] = entryFullDate
        viewToDate.send_keys(item['entryFullDate'])

        # Enter Reason for Absence
        viewAbsenceReason.send_keys(item['purpose'])

        # Complete Absence record
        btnAddAbsence.click()

        time.sleep(3)

        # Wait for the page to load before continuing
        mWait.until(EC.visibility_of_element_located((By.ID, "absenDestination")))

        # Disabled for debugging
        # btnSaveRecord.click()

finally:
    # Close the browser
    mDriver.quit()

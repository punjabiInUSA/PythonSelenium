import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait


"""
This script is designed to assist individuals who need to enter absence details for the physical presence calculator
 in the Canadian Citizenship portal (link provided below). Anyone can use it free of charge.

Link: https://citapply-citdemande.apps.cic.gc.ca/en/sign-in
"""


# Add file path
# data_only marked as true so only values are fetched even if those are generated from formulas
excelWorkbook = openpyxl.load_workbook("PATH_TO_EXCEL_FILE_WITH_FILENAME_AND_EXTENSION", data_only=True)
# Excel Format expected as
# Country | Exit Date | Exit Month | Exit Year | Entry Date | Entry Month | Entry Year | Purpose  | Exit Full Date | Entry Full Date - Header Row
# Brazil  | 15        | April      | 2023      | 24         | May         | 2023       | Vacation | 4/15/2023      | 5/24/2023 - Data Row

# Add sheet name (assumes excel has only one sheet)
# else use excelWorkbook['SPECIFY_SHEET_NAME']
targetSheetName = excelWorkbook.active

# Stores excel data
data = []

# Remove max_row=XX, if pulling complete data from excel
for row in targetSheetName.iter_rows(min_row=2, max_row=5, values_only=True):
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
    # print(item['exitDate'])
    # print(item['entryDate'])
    data.append(item)

# Create a new instance of Edge webdriver
mDriver = webdriver.Edge()
mWait = WebDriverWait(mDriver, 20)
webAddress = "https://citapply-citdemande.apps.cic.gc.ca/en/sign-in"
username = "ENTER_USER_ID"
password = "ENTER_PASSWORD"

# Navigate to the website
mDriver.get(webAddress)
userField = mDriver.find_element(By.ID, "emailAddressInput")
pwdField = mDriver.find_element(By.ID, "passwordInput")
btnSignIn = mDriver.find_element(By.ID, "signInButton")
userField.send_keys(username)
pwdField.send_keys(password)
btnSignIn.click()
time.sleep(2)

continueAppBtn = mWait.until(EC.
                             visibility_of_element_located((By.XPATH, "/html/body/app-root/app-shell/div[2]/div/"
                                                                      "div/div[1]/jl-cit-your-account/div/div/div[2]"
                                                                      "/div[2]/div/jl-cit-your-account-applicant-card/"
                                                                      "jl-cit-card-template/div/div/div[2]/"
                                                                      "jl-cit-button[1]/button")))

if continueAppBtn.is_displayed():
    # click retrieve saved calc button as website displays an error first time, requires re-initialization
    continueAppBtn.click()

physicalPresenceSection = mWait.until(
    EC.visibility_of_element_located((By.XPATH, "/html/body/app-root/app-shell/jl-cit-nav-bar/"
                                                "div/div[3]/div/div/div/ul/li[8]/a/span")))
physicalPresenceSection.click()
time.sleep(2)

# Question 5 about outside Canada trip details section
try:
    # will only process one row per argument data[:1], remove this to process all rows.
    for item in data:
        # Page Element Definitions
        viewAbsenceDestination = mDriver.find_element(By.ID, "country")
        fromYear = mDriver.find_element(By.ID, "date-exitYear-date-left-canada-default")
        fromMonth = mDriver.find_element(By.ID, "date-exitMonth-date-left-canada-default")
        fromDate = mDriver.find_element(By.ID, "date-exitDay-date-left-canada-default")
        toYear = mDriver.find_element(By.ID, "date-returnYear-date-return-to-canada-default")
        toMonth = mDriver.find_element(By.ID, "date-returnMonth-date-return-to-canada-default")
        toDate = mDriver.find_element(By.ID, "date-returnDay-date-return-to-canada-default")
        absenceReason = mDriver.find_element(By.ID, "absenceReasonSelectordefault")
        btnSaveRecord = mDriver.find_element(By.XPATH, "//*[@id=\"absenceFromCanadaSaveCardButtondefault\"]/button")

        # Focus onto absence entry area
        mDriver.execute_script("arguments[0].scrollIntoView();", viewAbsenceDestination)

        viewAbsenceDestination.send_keys(item['country'])

        fromYear.send_keys(item['exitYear'])
        fromMonth.send_keys(item['exitMonth'])
        fromDate.send_keys(item['exitDate'])

        toYear.send_keys(item['entryYear'])
        toMonth.send_keys(item['entryMonth'])
        toDate.send_keys(item['entryDate'])

        selectOtherAbsence = Select(absenceReason)
        selectOtherAbsence.select_by_visible_text("Other")
        absenceReasonDetails = mDriver.find_element(By.ID, "absenceFromCanadaFormDescription-otherReason")
        mDriver.execute_script("arguments[0].scrollIntoView();", absenceReason)
        absenceReasonDetails.send_keys(item['purpose'])
        time.sleep(1)

        btnSaveRecord.click()
        time.sleep(2)

finally:
    # Close the browser
    mDriver.quit()

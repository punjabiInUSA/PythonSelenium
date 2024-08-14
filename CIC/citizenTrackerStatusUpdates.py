import os.path
import sys
import time
import smtplib
from email.mime.text import MIMEText
import xlwings as xw
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from datetime import datetime
from selenium.webdriver import Edge, EdgeOptions

# Enable this if want to run headless, no visual UI
# options = Options()
# options.add_argument("--headless")
# options.add_argument("--window-size=1920,1080")
# driver = webdriver.Edge(options=options)

#Enable this for UI based execution
driver = webdriver.Edge()

mWait = WebDriverWait(driver, 30)

start_time = time.time()

wb = None
sheet = None

#Path and name for the excel file with updates
file_path = "C:\\Users\\username\\Downloads\\CitiTrackerUpdates.xlsx"

def initialize():
    # Navigate to Jenkins and activate credentials
    trackerUrl = f'https://tracker-suivi.apps.cic.gc.ca/en/login'
    driver.get(trackerUrl)
    usernameField = mWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="uci"]')))
    passwordField = mWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="password"]')))
    loginButton = mWait.until(EC.visibility_of_element_located((By.XPATH, '//button[contains(@class, "btn-sign-in")]')))
    # loginButton = mWait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/app-root/app-loading/div/app-shell/main/app-login/div/div/div[2]/form/button')))
    user = 'ENTER_USERNAME'
    pwd = 'ENTER_PASSWORD'
    if user == 'ENTER_USERNAME' or user == '':
        print('ERROR: Username not defined, review initialize() function')
        driver.quit()
        sys.exit()
    else:
        usernameField.send_keys(user)
    if pwd == 'ENTER_PASSWORD' or pwd == '':
        print('ERROR: Password not defined, review initialize() function')
        driver.quit()
        sys.exit()
    else:
        passwordField.send_keys(pwd)
        if loginButton.is_displayed():
            driver.execute_script("arguments[0].scrollIntoView();", loginButton)
            time.sleep(2)
        loginButton.click()

# Function extracts last update date, and status for each of the citizenship phases
# and enters the details into the excel sheet
def extract_info(sheet):
    # Set the headers for the columns
    sheet.range("A1").value = "Phase"
    sheet.range("B1").value = "Status"
    sheet.range("C1").value = "Last Updated"
    sheet.range("D1").value = "Last Checked"
    last_update = extract_last_update_date()
    current_date = datetime.now().strftime('%B %d, %Y')

    # Determine the starting row index for the data
    row_index = sheet.range("A1").current_region.last_cell.row + 1 if sheet.range("A2").value else 2
    print(f'Row index: {row_index}')
    print('Current Date: ' + current_date)

    #Captures last update date into excel
    sheet.range(f'C{row_index}').value = last_update

    # Captures script execution date into excel
    sheet.range(f'D{row_index}').value = current_date

    # Locate the details-section and then find all li elements within that section
    details_section = driver.find_element(By.XPATH, '//section[contains(@class, "details-section")]')
    li_elements = details_section.find_elements(By.XPATH, './/ul//li')

    for li in li_elements:
        # Extract the Phase name
        h3_text = li.find_element(By.XPATH, './/h3').text

        # Extract the phase status
        chip_text_element = li.find_element(By.XPATH, './/p[contains(@class, "chip-text")]')
        chip_text = chip_text_element.text

        # Capture phase name and status into Excel
        sheet.range(f'A{row_index}').value = h3_text
        sheet.range(f'B{row_index}').value = chip_text

        # Move to the next row in Excel
        row_index += 1

#Fetches last updated date
def extract_last_update_date():
    date_element = mWait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "dl.date-container dd.date-text")))
    last_update_date = date_element.text
    print('Last Updated:'+ last_update_date)
    return last_update_date


def create_workbook():
    global wb, sheet
    if os.path.exists(file_path):
        # Create a new Excel workbook
        print(f"Opening existing workbook: {file_path}")
        wb = xw.Book(file_path)
        # Activate the first sheet in the workbook
        sheet = wb.sheets[0]
    else:
        print(f"Creating new workbook at: {file_path}")
        # Create a new workbook if it doesn't exist
        wb = xw.Book()
        sheet = wb.sheets[0]
    return wb, sheet


def save_workbook(workbook):
    global wb
    print('Saving file to: ' + file_path)
    # Save the workbook
    workbook.save(file_path)
    workbook.close()

#Script execution starts here
initialize()
wb, sheet = create_workbook()
extract_info(sheet)
save_workbook(wb)

# Close the browser
driver.quit()

end_time = time.time()
elapsed_time = end_time - start_time
print(f'Script completed, total execution time: {elapsed_time} seconds')



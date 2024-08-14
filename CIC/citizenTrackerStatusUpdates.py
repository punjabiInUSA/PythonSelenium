import os.path
import sys
import time
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Enable this if want to run headless, no visual UI
# options = Options()
# options.add_argument("--headless")
# options.add_argument("--window-size=1920,1080")
# driver = webdriver.Edge(options=options)

# Enable this for UI based execution
driver = webdriver.Edge()

mWait = WebDriverWait(driver, 30)

start_time = time.time()

# Path and name for the text file with updates
file_path = "C:\\Users\\username\\Downloads\\CitiTrackerUpdates.txt"

def initialize():
    # Navigate to Jenkins and activate credentials
    trackerUrl = f'https://tracker-suivi.apps.cic.gc.ca/en/login'
    driver.get(trackerUrl)
    usernameField = mWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="uci"]')))
    passwordField = mWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="password"]')))
    loginButton = mWait.until(EC.visibility_of_element_located((By.XPATH, '//button[contains(@class, "btn-sign-in")]')))
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
# and enters the details into the text file
def extract_info(file):
    last_update = extract_last_update_date()
    current_date = datetime.now().strftime('%B %d, %Y')

    # Write headers if file is empty
    if os.stat(file_path).st_size == 0:
        file.write("Phase\tStatus\n")

    # Locate the details-section and then find all li elements within that section
    details_section = driver.find_element(By.XPATH, '//section[contains(@class, "details-section")]')
    li_elements = details_section.find_elements(By.XPATH, './/ul//li')

    for li in li_elements:
        # Extract the Phase name
        h3_text = li.find_element(By.XPATH, './/h3').text

        # Extract the phase status
        chip_text_element = li.find_element(By.XPATH, './/p[contains(@class, "chip-text")]')
        chip_text = chip_text_element.text

        # Write phase name and status into the text file
        file.write(f"{h3_text}\t{chip_text}\n")

    # Write last update and current date on separate lines
    file.write(f"Last Updated:\t {last_update}\n")
    file.write(f"Script Date:\t {current_date}\n\n\n")

# Fetches last updated date
def extract_last_update_date():
    date_element = mWait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "dl.date-container dd.date-text")))
    last_update_date = date_element.text
    print('Last Updated:' + last_update_date)
    return last_update_date

# Script execution starts here
initialize()

# Open the text file in append mode
with open(file_path, 'a') as file:
    extract_info(file)

# Close the browser
driver.quit()

end_time = time.time()
elapsed_time = end_time - start_time
print(f'Script completed, total execution time: {elapsed_time} seconds')
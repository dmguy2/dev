from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import csv
import datetime
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import time


# initialize the Chrome driver
driver = webdriver.Chrome("chromedriver")
driver.maximize_window()
driver.implicitly_wait(30)

# Github credentials
username, password = "david.guy", "EdIson3.20"

# Scrape Limit
scrape_limit = 101

# head to shelby login page
driver.get("https://groverva.shelbynextchms.com/user/login")
driver.find_element(By.ID, "login-username").send_keys(username)
driver.find_element(By.ID, "login-password").send_keys(password)
driver.find_element(By.ID, "btn-login").click()

# navigate to attendance page and show scrape_limit records per page
driver.get("https://groverva.shelbynextchms.com/reports/attendance/sessions")
wait = WebDriverWait(driver, 30)
wait.until(EC.presence_of_element_located((By.ID, "react-panel-view")))
select = Select(driver.find_element(
    By.XPATH, '//*[@id="react-panel-view"]/span/div[3]/span[1]/span/select'))
select.select_by_value("500")

# scrape table and export data to csv
wait.until(EC.presence_of_element_located(
    (By.XPATH, '//*[@id="react-panel-view"]/span/div[2]/table/tbody/tr[1]')))
page_source = driver.page_source
soup = BeautifulSoup(page_source, "html.parser")
current_date = datetime.datetime.now().strftime("%Y-%m-%d")
filename = f"attendance_{current_date}.csv"
with open(filename, mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["Session", "Date", "Present", "Reason",
                    "No Reason", "Visitors", "% Accounted For"])
    for i in range(1, scrape_limit):
        xpath = f'//*[@id="react-panel-view"]/span/div[2]/table/tbody/tr[{i}]'
        row = driver.find_element(By.XPATH, xpath)
        cells = row.find_elements(By.XPATH, f"{xpath}/td")
        writer.writerow([cell.text for cell in cells])

# append data to google sheets
page_source = driver.page_source
soup = BeautifulSoup(page_source, "html.parser")
SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/drive.file']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'credentials.json', SCOPES)
client = gspread.authorize(creds)
worksheet = client.open("Grove Ops Weekly Snapshot").worksheet(
    "Planning Center Scrape")
for i in range(1, scrape_limit):
    time.sleep(1)
    xpath = f'//*[@id="react-panel-view"]/span/div[2]/table/tbody/tr[{i}]'
    row = driver.find_element(By.XPATH, xpath)
    cells = row.find_elements(By.XPATH, f"{xpath}/td")
    worksheet.append_row([cell.text for cell in cells])

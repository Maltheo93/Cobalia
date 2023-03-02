import calendar
import csv
import glob
import locale
import os
import re
import shutil
import time
from datetime import datetime, timedelta, date
import openpyxl
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

# Set up the driver with notifications disabled, downloads allowed and then navigate to Cobália
options = Options()
options.add_argument("--disable-notifications")
options.add_argument("--safebrowsing-disable-download-protection")
prefs = {
    "download.default_directory": os.path.join(os.path.expanduser("~"), "Downloads"),
    "download.prompt_for_download": False,
    "safebrowsing.enabled": False,
    "safebrowsing.disable_download_protection": True,
    "profile.default_content_settings.popups": 0,
    "profile.default_content_setting_values.automatic_downloads": 1,
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)
driver.maximize_window()
driver.get("https://www.cobalia.com")

# Locate the Email & Password fields and enter information and then locating and clicking login button
filepath = r'C:\Users\mamo\Desktop\Dambrug Monthly Reports\Login Information.txt'
with open(filepath, 'r') as f:
    lines = f.readlines()
    username = lines[0].strip().split(':')[1].strip()
    password = lines[1].strip().split(':')[1].strip()
email_field = WebDriverWait(driver, 10).until(
    ec.element_to_be_clickable((By.CSS_SELECTOR, 'input[cy-data-id="input-username"]')))
email_field.send_keys(username)
password_field = driver.find_element(By.CSS_SELECTOR, 'input[cy-data-id="input-password"]')
password_field.send_keys(password)
driver.find_element(By.CSS_SELECTOR, '[cy-data-id="button-login"]').click()

# Go to Bryrup Dambrug
input_elements = ['[cy-data-id="btn-unit-dropdown"]', '[class="header"]',
                  '[title="#41"]']
for input_ in input_elements:
    element = WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, input_)))
    element.click()

# Go to Facility Report
time.sleep(0.5)
input_elements = ['/html/body/cb-app-root/app-home/cb-side-nav/nav/ul/li[6]/a',
                  '/html/body/cb-app-root/app-home/cb-side-nav/nav/ul/li[6]/div/ul/li[2]/div/a',
                  '//*[@id="sidebar"]/ul/li[6]/div/ul/li[2]/ul/li[1]/a',
                  '/html/body/cb-app-root/app-home/cb-side-nav/nav/ul/li[6]/a']
for input_ in input_elements:
    element = WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.XPATH, input_)))
    element.click()

# Set date range to 'User Defined'
time.sleep(2)
input_elements = [
    '//*[@id="container"]/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/cb-tb-select',
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/cb-tb-select/div/div[2]/cb-tb-option[4]',
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/tb-flatpickr[1]/div/input']
for input_ in input_elements:
    element = WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.XPATH, input_)))
    element.click()

# In Calendar 1, proceed to previous month and click the first day that does not have labels for previous or next month
input_elements = [
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/tb-flatpickr[1]/div/input',
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/tb-flatpickr[1]/div/div/div[1]/span[1]',
    '//span[@class="flatpickr-day" and @aria-label and not(contains(@class, "prevMonthDay") or contains(@class, "nextMonthDay"))][1]']
for input_ in input_elements:
    element = WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.XPATH, input_)))
    element.click()

# In Calendar 2, proceed to previous month and click the last day of the previous month
input_elements = [
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/tb-flatpickr[2]/div/input',
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/tb-flatpickr[2]/div/div/div[1]/span[1]']
for input_ in input_elements:
    element = WebDriverWait(driver, 5).until(ec.element_to_be_clickable((By.XPATH, input_)))
    element.click()
for i in range(31, 27, -1):
    day_xpath = f'//span[@class="flatpickr-day" and @aria-label and not(contains(@class, "prevMonthDay") or contains(@class, "nextMonthDay"))][{i}]'
    try:
        day_element = WebDriverWait(driver, 0.5).until(
            ec.element_to_be_clickable((By.XPATH, day_xpath)))
        day_element.click()
        break
    except TimeoutException:
        if i == 28:
            raise ValueError("Unable to find the last day of the month.")
        continue

# Wait for Data to load and Download PDF of Facility Report
time.sleep(3)
input_elements = [
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/cb-tb-right-align/cb-tb-dropdown',
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/cb-tb-right-align/cb-tb-dropdown/div[2]/cb-tb-dropdown-item[1]',
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-print-configuration/ui-modal/div/div/div/div/button[1]',
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-print-configuration/ui-modal/div/div/button']
for input_ in input_elements:
    element = WebDriverWait(driver, 5).until(ec.element_to_be_clickable((By.XPATH, input_)))
    element.click()

# Download CSV of Facility Report
time.sleep(0.5)
input_elements = [
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/cb-tb-right-align/cb-tb-dropdown',
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-standard-layout/cb-content/div/cb-tb-toolbar/cb-tb-right-align/cb-tb-dropdown/div[2]/cb-tb-dropdown-item[2]',
    '/html/body/cb-app-root/app-home/div/div/facility-report-v2/cb-print-configuration/ui-modal/div/div/button']
for input_ in input_elements:
    element = WebDriverWait(driver, 5).until(ec.element_to_be_clickable((By.XPATH, input_)))
    element.click()

# Delete the newest file in the directory that contains "Facilitetsrapport" in the filename and has a .csv extension
time.sleep(3)
download_path = "C:/Users/mamo/Downloads"
file_pattern = os.path.join(download_path, "*Facilitetsrapport*.csv")
files = glob.glob(file_pattern)
if len(files) > 0:
    newest_file = max(files, key=os.path.getctime)
    os.remove(newest_file)

# Calculate previous month and specify source and destination folders
time.sleep(0.5)
now = datetime.now()
previous_month = now.replace(day=1) - timedelta(days=1)
previous_month_str = previous_month.strftime("%m-%Y")
bryrup_path = "C:/Users/mamo/Desktop/Dambrug Monthly Reports/Bryrup"

# get the list of csv and pdf files in the source folder, sorted by modification time
csv_files = glob.glob(os.path.join(download_path, "*.csv"))
pdf_files = glob.glob(os.path.join(download_path, "*.pdf"))
csv_files.sort(key=os.path.getmtime)
pdf_files.sort(key=os.path.getmtime)

# get the name of the newest csv and pdf files, and construct new names before moving to destination
newest_csv_file = csv_files[-1]
newest_pdf_file = pdf_files[-1]
new_csv_file_name = f"{previous_month_str} Bryrup.csv"
new_pdf_file_name = f"{previous_month_str} Bryrup.pdf"
shutil.move(newest_csv_file, os.path.join(bryrup_path, new_csv_file_name))
shutil.move(newest_pdf_file, os.path.join(bryrup_path, new_pdf_file_name))

# Set up paths and names, then copy and rename file
dambrug_path = r"C:\Users\mamo\Desktop\Dambrug Monthly Reports"
source_file_name = "Skabelon.xlsx"
today = date.today()
prev_month = today.replace(day=1) - timedelta(days=1)
dest_file_name = "{:02d}-{} Bryrup.xlsx".format(prev_month.month, prev_month.year)
shutil.copy(os.path.join(dambrug_path, source_file_name), bryrup_path)
os.rename(os.path.join(bryrup_path, source_file_name), os.path.join(bryrup_path, dest_file_name))

# Set up file paths and names
folder_name = os.path.basename(bryrup_path)
previous_month = (date.today().replace(day=1) - timedelta(days=1)).replace(day=1)
prev_month_name = calendar.month_name[previous_month.month][:3]
prev_month_name_da = \
    "Januar Februar Marts April Maj Juni Juli August September Oktober November December".split()[
        previous_month.month - 1]
prev_month_year = previous_month.year

# Load the XLSX file and get a reference to the existing sheet
destination_file = os.path.join(bryrup_path, f'{previous_month.strftime("%m-%Y")} {folder_name}.xlsx')
book = openpyxl.load_workbook(destination_file)
sheet = book['Data']

# Read the data from the CSV file into a list of lists
with open(
        os.path.join(bryrup_path, f'{previous_month.strftime("%m-%Y")} {folder_name}.csv'),
        newline='', encoding='utf-8-sig') as csvfile:
    reader = csv.reader(csvfile, delimiter='	')
    data = list(reader)

# Set the locale to handle Danish number system
locale.setlocale(locale.LC_ALL, 'da_DK.UTF-8')

# Loop through the list of lists and write the data to the sheet
for row_index, row_data in enumerate(data):
    for col_index, col_data in enumerate(row_data):
        # Check if the data is a number
        if re.match(r'^[0-9,.]+$', col_data):
            # Convert the string to a float using the Danish number system
            col_data = locale.atof(col_data)
            # Set the cell format to be numeric
            sheet.cell(row=row_index + 4, column=col_index + 1).number_format = '#,##0.00'
        sheet.cell(row=row_index + 4, column=col_index + 1).value = col_data

# Update Summation sheet
summation_sheet = book['Summation']
summation_sheet['C4'] = folder_name
summation_sheet['C6'] = prev_month_name_da
summation_sheet['C8'] = prev_month_year

book.save(destination_file)
book.close()

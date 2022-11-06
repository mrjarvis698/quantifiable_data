from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.support.ui import Select
import os
from os import path
import json, shutil, time
from tkinter import filedialog
import pandas as pd
from openpyxl import load_workbook
from openpyxl.workbook import Workbook

# Open xlsx file
open_sheet = path.exists("temp/opened_sheet.json")
if open_sheet == True:
    
    opened_sheet_file_path = "temp/opened_sheet.json"
    json_file = open(opened_sheet_file_path)
    data = json.load(json_file)
    xlsx_sheet_check = path.exists(data ['xlsx_file_path'])
    if xlsx_sheet_check == True :
        xlsx_file_path = data['xlsx_file_path']
        
    else :
        shutil.rmtree('temp', ignore_errors=True)
        xlsx_file_path = filedialog.askopenfilename(title="Open Excel-XLSX File")
        cache_path = os.path.join(str(os.getcwd()), "temp")
        dictionary = {"xlsx_file_path" : xlsx_file_path}
        json_object = json.dumps(dictionary, indent = 1)
        with open("temp/opened_sheet.json", "w") as outfile:
            outfile.write(json_object)
else :
    xlsx_file_path = filedialog.askopenfilename(title="Open Excel-XLSX File")
    cache_path = os.path.join(str(os.getcwd()), "temp")
    os.mkdir(cache_path)
    dictionary = {"xlsx_file_path" : xlsx_file_path}
    json_object = json.dumps(dictionary, indent = 1)
    with open("temp/opened_sheet.json", "w") as outfile:
        outfile.write(json_object)

# read imported xlsx file path using pandas
input_workbook = pd.read_excel(xlsx_file_path, sheet_name = 'Sheet1', dtype=str)
total_input_rows, total_input_cols = input_workbook.shape

input_col = list(input_workbook.columns.values.tolist())

srnumber = input_workbook[input_col[0]].values.tolist()
mobile = input_workbook[input_col[1]].values.tolist()
name_in_hindi = input_workbook[input_col[2]].values.tolist()
name_in_English = input_workbook[input_col[3]].values.tolist()
father_name_in_hindi = input_workbook[input_col[4]].values.tolist()
father_name_in_english = input_workbook[input_col[5]].values.tolist()
gender = input_workbook[input_col[6]].values.tolist()
address = input_workbook[input_col[7]].values.tolist()
district = input_workbook[input_col[8]].values.tolist()

# get-output sheet to append output
output_sheet = path.exists("Output.xlsx")
if output_sheet == True :
    output_sheet_file_path = "Output.xlsx"
else :
    input_col.append('Status')
    output_headers = input_col
    overall_output = Workbook()
    page = overall_output.active
    page.append(output_headers)
    overall_output.save(filename = 'Output.xlsx')
    output_sheet_file_path = "Output.xlsx"

def test_basic_options():
    shutil.rmtree('C:/Users/Sairaj/.wdm/drivers/chromedriver/', ignore_errors=True)
    global driver
    service = ChromeService(executable_path=ChromeDriverManager().install())
    chrome_options = ChromeOptions()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-crash-reporter")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-in-process-stack-traces")
    chrome_options.add_argument("--disable-logging")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--output=/dev/null")
    driver = webdriver.Chrome(options=chrome_options, service=service)
    driver.implicitly_wait(4)
    driver.maximize_window()
    
def script_cmd(skip):
    driver.get("https://cgqdc.in/cgqdc-user-registration")
    
    if skip == False:
        try:
            driver.find_element(By.ID,"mobile_btn")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"mobile_btn").click()

    if skip == False:
        try:
            driver.find_element(By.ID,"han")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"han").click()

    if skip == False:
        try:
            driver.find_element(By.ID,"mobile_no")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"mobile_no").send_keys(mobile[i])

    if skip == False:
        try:
            driver.find_element(By.ID,"valiIpt")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"valiIpt").send_keys(input("Enter Captcha : "))

    if skip == False:
        try:
            driver.find_element(By.ID,"mobile_send_otp")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"mobile_send_otp").click()

    if skip == False:
        try:
            driver.find_element(By.ID,"otp")
        except NoSuchElementException:
            try:
                driver.find_element(By.XPATH,"(.//*[normalize-space(text()) and normalize-space(.)='आप पहले से ही पंजीकृत है|'])[1]/following::button[1]")
            except:
                skip = True
                skip_reason = "Failed"
            else:
                driver.find_element(By.XPATH,"(.//*[normalize-space(text()) and normalize-space(.)='आप पहले से ही पंजीकृत है|'])[1]/following::button[1]").click()
                skip = True
                skip_reason = "Already Registered"
        else:
            skip = False
            driver.find_element(By.ID,"otp").send_keys("654321")

    if skip == False:
        try:
            driver.find_element(By.ID,"submit")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"submit").click()

    if skip == False:
        try:
            driver.find_element(By.ID,"mukhiya_name_hindi_1")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"mukhiya_name_hindi_1").send_keys(name_in_hindi[i])

    if skip == False:
        try:
            driver.find_element(By.ID,"mukhiya_name_eng_1")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"mukhiya_name_eng_1").send_keys(name_in_English[i])

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_fh_name_hin")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"applicant_fh_name_hin").send_keys(father_name_in_hindi[i])

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_fh_name_eng")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"applicant_fh_name_eng").send_keys(father_name_in_english[i])

    if skip == False:
        try:
            driver.find_element(By.ID,"mukhiya_gender_1")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            try :
                Select(driver.find_element(By.ID,"mukhiya_gender_1")).select_by_value(gender[i])
            except ElementNotInteractableException:
                time.sleep(1)
                try:
                    Select(driver.find_element(By.ID,"mukhiya_gender_1")).select_by_value(gender[i])
                except ElementNotInteractableException:
                    skip = True
                    skip_reason = "Failed"
                else :
                    skip = False
            else:
                skip = False

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_caste")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            try :
                Select(driver.find_element(By.ID,"applicant_caste")).select_by_value("3_अन्य पिछड़ा वर्ग")
            except ElementNotInteractableException:
                time.sleep(1)
                try:
                    Select(driver.find_element(By.ID,"applicant_caste")).select_by_value("3_अन्य पिछड़ा वर्ग")
                except ElementNotInteractableException:
                    skip = True
                    skip_reason = "Failed"
                else :
                    skip = False
            else:
                skip = False

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_caste_category")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            try :
                Select(driver.find_element(By.ID,"applicant_caste_category")).select_by_value("51")
            except ElementNotInteractableException:
                time.sleep(1)
                try:
                    Select(driver.find_element(By.ID,"applicant_caste_category")).select_by_value("51")
                except ElementNotInteractableException:
                    skip = True
                    skip_reason = "Failed"
                else :
                    skip = False
            else:
                skip = False

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_district")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            try :
                Select(driver.find_element(By.ID,"applicant_district")).select_by_value("49_646_बालोद ")
            except ElementNotInteractableException:
                time.sleep(1)
                try:
                    Select(driver.find_element(By.ID,"applicant_district")).select_by_value("49_646_बालोद ")
                except ElementNotInteractableException:
                    skip = True
                    skip_reason = "Failed"
                else :
                    skip = False
            else:
                skip = False

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_area")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            try :
                Select(driver.find_element(By.ID,"applicant_area")).select_by_value("V")
            except ElementNotInteractableException:
                time.sleep(1)
                try:
                    Select(driver.find_element(By.ID,"applicant_area")).select_by_value("V")
                except ElementNotInteractableException:
                    skip = True
                    skip_reason = "Failed"
                else :
                    skip = False
            else:
                skip = False

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_block")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            try :
                Select(driver.find_element(By.ID,"applicant_block")).select_by_value("5_3629_बालौद")
            except ElementNotInteractableException:
                time.sleep(1)
                try:
                    Select(driver.find_element(By.ID,"applicant_block")).select_by_value("5_3629_बालौद")
                except ElementNotInteractableException:
                    skip = True
                    skip_reason = "Failed"
                else :
                    skip = False
            else:
                skip = False

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_gram_panchayat")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            try :
                Select(driver.find_element(By.ID,"applicant_gram_panchayat")).select_by_value("124065_43050390_दरबारी नवागांव")
            except ElementNotInteractableException:
                time.sleep(1)
                try:
                    Select(driver.find_element(By.ID,"applicant_gram_panchayat")).select_by_value("124065_43050390_दरबारी नवागांव")
                except ElementNotInteractableException:
                    skip = True
                    skip_reason = "Failed"
                else :
                    skip = False
            else:
                skip = False
    
    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_gram")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            try :
                Select(driver.find_element(By.ID,"applicant_gram")).select_by_value("3946_द.नवागांव_443147")
            except ElementNotInteractableException:
                time.sleep(1)
                try:
                    Select(driver.find_element(By.ID,"applicant_gram")).select_by_value("3946_द.नवागांव_443147")
                except ElementNotInteractableException:
                    skip = True
                    skip_reason = "Failed"
                else :
                    skip = False
            else:
                skip = False

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_pincode")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"applicant_pincode").send_keys("400001")

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_address")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"applicant_address").send_keys(address[i])

    if skip == False:
        try:
            driver.find_element(By.ID,"applicant_member_work")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            try :
                Select(driver.find_element(By.ID,"applicant_member_work")).select_by_value("1")
            except ElementNotInteractableException:
                time.sleep(1)
                try:
                    Select(driver.find_element(By.ID,"applicant_member_work")).select_by_value("1")
                except ElementNotInteractableException:
                    skip = True
                    skip_reason = "Failed"
                else :
                    skip = False
            else:
                skip = False

    if skip == False:
        try:
            driver.find_element(By.XPATH,"//form[@id='fromTarget']/div/div[24]/div/label")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.XPATH,"//form[@id='fromTarget']/div/div[24]/div/label").click()

    if skip == False:
        try:
            driver.find_element(By.ID,"date")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            input("Enter Date of birth and Press Enter Key to continue.")

    if skip == False:
        try:
            driver.find_element(By.ID,"formSubmit")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.ID,"formSubmit").click()
    
    if skip == False:
        try:
            driver.find_element(By.XPATH,"(.//*[normalize-space(text()) and normalize-space(.)='चेक करें'])[1]/following::button[1]")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            driver.find_element(By.XPATH,"(.//*[normalize-space(text()) and normalize-space(.)='चेक करें'])[1]/following::button[1]").click()

    if skip == False:
        try:
            driver.find_element(By.XPATH,"//button[@type='button']")
        except NoSuchElementException:
            skip = True
            skip_reason = "Failed"
        else:
            skip = False
            skip_reason = "Success"
            time.sleep(10)

    output = srnumber[i], mobile[i], name_in_English[i],gender[i], skip_reason
    print(output)

for i in range (0,total_input_rows):
    test_basic_options()
    script_cmd(False)

driver.quit()
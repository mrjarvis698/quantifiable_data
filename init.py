from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time

def test_basic_options():
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
    driver.maximize_window()
    driver.get("https://cgqdc.in/")
    driver.find_element(By.LINK_TEXT,"पंजीयन").click()
    driver.find_element(By.ID,"mobile_btn").click()
    driver.find_element(By.ID,"han").click()
    time.sleep(5)
    driver.find_element(By.ID,"mobile_no").click()
    driver.find_element(By.ID,"mobile_no").send_keys("8103619402")
    time.sleep(10)
    driver.find_element(By.ID,"mobile_send_otp").click()
    driver.quit()

test_basic_options()
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
PATH = r"C:\Users\YuChen\Desktop\chromedriver"
EMAILFIELD = (By.ID, "i0116")
PASSWORDFIELD = (By.ID, "i0118")
NEXTBUTTON = (By.ID, "idSIButton9")
username = input("Enter Your Username: ")

password = input("Enter Your Password: ")

browser = webdriver.Chrome(PATH)
browser.get('https://login.live.com')

# wait for email field and enter email
WebDriverWait(browser, 10).until(EC.element_to_be_clickable(EMAILFIELD)).send_keys(username)

# Click Next
WebDriverWait(browser, 10).until(EC.element_to_be_clickable(NEXTBUTTON)).click()

# wait for password field and enter password
WebDriverWait(browser, 10).until(EC.element_to_be_clickable(PASSWORDFIELD)).send_keys(password)

# Click Login - same id?
WebDriverWait(browser, 10).until(EC.element_to_be_clickable(NEXTBUTTON)).click()

#go to garena folder (change garena folder to alerts folder)
browser.get("https://outlook.live.com/mail/0/AQMkADAwATY3ZmYAZS1hYmQ2LWQzYjQtMDACLTAwCgAuAAADsccrHdzPa02bB4CVQ7k9uAEAjZADpJtFskWUXgzqucJvvQAAAgFPAAAA/id/AQMkADAwATY3ZmYAZS1hYmQ2LWQzYjQtMDACLTAwCgBGAAADsccrHdzPa02bB4CVQ7k9uAcAjZADpJtFskWUXgzqucJvvQAAAgFPAAAAjZADpJtFskWUXgzqucJvvQAAAl3IAAAA");

emails = browser.find_elements_by_class_name("YWkvAfVxfWoYoGc_xj-4c")
for email in emails:
    print(email.text)
    test = (By.CLASS_NAME, "YWkvAfVxfWoYoGc_xj-4c")
    WebDriverWait(browser,10).until(EC.element_to_be_clickable(test))
    email.click()
    time.sleep(1)
    browser.back()
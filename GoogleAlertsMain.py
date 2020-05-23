from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import csv
from selenium.webdriver.common.keys import Keys

PATH = r"C:\Users\YuChen\Desktop\chromedriver"

print("Please enter Keyword: ")
keyword = input()

print("Please enter the integer")
print("1: Append the file")
print("2: Overwrite the file")

whattodo = int(input())
if whattodo == 1:
    alpha = 'a'
elif whattodo == 2:
    alpha = 'w'

browser = webdriver.Chrome(PATH)
SEARCH = (By.CLASS_NAME, "label-input-label")

browser.get('https://www.google.com.sg/alerts#')

# wait for page to load
element = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div/div[1]/div[1]/div/input"))
    )
element.send_keys(keyword)
# wait for page to load
time.sleep(2)
# get results and sources
links = browser.find_elements_by_class_name("result_title_link")
sources = browser.find_elements_by_class_name("result_source")

# variable for getting sources using list
i = 0

# open csv file
with open('mycsv.csv', alpha, newline='') as f:
    fieldnames = ['Keyword', 'Title', 'Link', 'Source']
    thewriter = csv.DictWriter(f, fieldnames=fieldnames)
    if whattodo == 2:
        thewriter.writeheader()
    # iterate through the list of results and writing in csv file
    for link in links:
        print(link.text)
        print(link.get_attribute("href"))
        print(sources[i].text)
        thewriter.writerow({'Keyword' : keyword, 'Title' : link.text, 'Link' : link.get_attribute("href") , 'Source' : sources[i].text})
        i = i + 1
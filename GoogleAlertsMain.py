from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
import csv
from selenium.webdriver.common.keys import Keys

# path where chromedriver is located
PATH = r"C:\Users\YuChen\Desktop\chromedriver"

# keyword for searching
print("Please enter Keyword: ")
keyword1 = input()

# search parameters
print("Would you like to refine the search? Y/N")
refineSearch = input()
if refineSearch == "Y" or refineSearch == "y":
    print("How would you like to refine the search?")
    print("1. Combine search term")
    print("2. Search social media")
    print("3. Exclude a word")
    print("4. Search within a website")
    searchOption = input()
    if searchOption == "1":
        print("Please enter second keyword:")
        keyword2 = input()
        keyword = keyword1 + " OR " + keyword2
    elif searchOption == "2":
        keyword = "@" + keyword1
    elif searchOption == "3":
        print("Please enter word to exclude:")
        keyword2 = input()
        keyword = keyword1 + " -" + keyword2
    elif searchOption == "4":
        print("Please enter website")
        keyword2 = input()
        keyword = keyword1 + " site:" + keyword2
else:
    keyword = keyword1


# let the csv writer know whether to append or write
print("Please enter the integer")
print("1: Append the file")
print("2: Overwrite the file")
whattodo = int(input())
if whattodo == 1:
    alpha = 'a'
elif whattodo == 2:
    alpha = 'w'



# open chrome
browser = webdriver.Chrome(PATH)

# go to website
browser.get('https://www.google.com.sg/alerts#')

# locate the search box and send keyword
element = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div/div[1]/div[1]/div/input"))
    )
element.send_keys(keyword)

# change results to all results
show_options = WebDriverWait(browser,10).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div/div[1]/div[2]/div[2]/div[2]/span[3]")))
show_options.click()

show = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div/div[1]/div[2]/div[2]/div[1]/table/tbody/tr[5]/td[2]/div/div[1]"))
    )
show.click()

show_expanded = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div"))
    )
show_expanded.click()

# wait for page to load
WebDriverWait(browser,10).until(EC.presence_of_element_located((By.CLASS_NAME, "result_title_link")))

# get results and sources in a list
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
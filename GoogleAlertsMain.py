from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.common.by import By
import csv
from datetime import datetime
from selenium.webdriver.common.keys import Keys
import xlsxwriter
from openpyxl import Workbook
import openpyxl
# path where chromedriver is located
# PATH = r":C:\Users\yc_19\Documents\chromedriver"
PATH = r"C:\Users\YuChen\Desktop\chromedriver"

# keyword for searching
print("Please enter Keyword: ")
keyword = input()

# search parameters
# 1. Combine search term (keyword1 AND keyword2)
# 2. Search social media (@keyword1)
# 3. Exclude a word (keyword -wordtoexclude)
# 4. Search within a website (keyword site:website)

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
show_options = WebDriverWait(browser, 10).until(
    EC.presence_of_element_located(
        (By.XPATH, "/html/body/div[2]/div[2]/div[1]/div/div[1]/div[2]/div[2]/div[2]/span[3]")))
show_options.click()

show = WebDriverWait(browser, 10).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div/div[1]/div[2]/div[2]/div[1]"
                                              "/table/tbody/tr[5]/td[2]/div/div[1]"))
)
show.click()

show_expanded = WebDriverWait(browser, 10).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div"))
)
show_expanded.click()

# wait for page to load
WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "result_title_link")))

# get results and sources in a list
links = browser.find_elements_by_class_name("result_title_link")
sources = browser.find_elements_by_class_name("result_source")

# variable for getting sources using list
i = 0
e1 = 0
e2 = 0
e3 = 0
s = 0
main_window = browser.current_window_handle
now = datetime.now()
date = now.strftime("%d-%m-%y %H%M")
print(date)

# open csv file
with open(keyword + ".csv", 'a', newline='') as f:
    fieldnames = ['Keyword', 'Title', 'Link', 'Source', 'Paragraph']
    thewriter = csv.DictWriter(f, fieldnames=fieldnames)
    thewriter.writeheader()
    try:
        wb = openpyxl.load_workbook(keyword + '.xlsx')
        wb.create_sheet(date)
        ws = wb[date]
    except Exception as e:
        print(e)
        wb = Workbook()
        ws = wb.active
        ws.title = date
    row = 2
    col = 0
    ws['A1'] = 'Keyword'
    ws['B1'] = 'Title'
    ws['C1'] = 'Link'
    ws['D1'] = 'Source'
    ws['E1'] = 'Paragraph'

    # iterate through the list of results and writing in csv file
    for link in links:
        try:
            print(link.text)
            print(link.get_attribute("href"))
            print(sources[i].text)
            # Use: Keys.CONTROL + Keys.SHIFT + Keys.RETURN to open tab on top of the stack

            link.send_keys(Keys.CONTROL + Keys.RETURN)
            # Switch to the new window and open URL B
            browser.switch_to.window(browser.window_handles[1])
            WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.TAG_NAME, "p")))
            paragraphs = browser.find_elements_by_tag_name("p")
            paragraph_summary = ""
            for paragraph in paragraphs:
                paragraph_summary = paragraph_summary + paragraph.text
                print(paragraph.text)
        except Exception as e:
            print(e)
            e1 += 1

        browser.close()
        browser.switch_to.window(browser.window_handles[0])

        try:
            thewriter.writerow(
                {'Keyword': keyword, 'Title': link.text, 'Link': link.get_attribute("href"), 'Source': sources[i].text,
                 'Paragraph': paragraph_summary})
        except Exception as e:
            print(e)
            e2 += 1
        try:
            ws.cell(row=row, column=1).value = keyword
            ws.cell(row=row, column=2).value = link.text
            ws.cell(row=row, column=3).value = link.get_attribute("href")
            ws.cell(row=row, column=4).value = sources[i].text
            ws.cell(row=row, column=5).value = paragraph_summary
            row += 1
            s += 1
        except Exception as e:
            print(e)
            e3 += 1
        i += 1
        print("success", s)
        print("paragraphError:", e1, "csvError:", e2, "xlsxError:", e3)
        if i == 5:
            break
wb.save(keyword + '.xlsx')
browser.close()

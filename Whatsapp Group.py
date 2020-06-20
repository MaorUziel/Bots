from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from pandas import DataFrame
import pandas as pd
import xlsxwriter

from datetime import datetime
from time import sleep
import random
from random import randint


driver = webdriver.Firefox(executable_path="C:\webdrivers\geckodriver.exe")
driver.get("https://web.whatsapp.com")


new_sheet = "Group" + str(datetime.now())
filename = r'contact.xlsx'
df = pd.read_excel(filename)
names = list(df['name'])

time_now = str(datetime.today())[0:16]
time_now_fixed = time_now.replace(":", ".")
status_sheet = "status " + time_now_fixed + ".xlsx"
workbook = xlsxwriter.Workbook(status_sheet)
sheet = workbook.add_worksheet("Status")


input("Enter anything after scanning QR code")

menu_bar = driver.find_element_by_xpath('//div[@role="button"][@title="תפריט"]')
menu_bar.click()
sleep(randint(1,2))
new_group = driver.find_element_by_xpath('//div[@role="button"][@title="קבוצה חדשה"]')
new_group.click()
sleep(randint(1,2))

row = 0
for i in names:
    try:
        contact_seacrch_box = driver.find_element_by_xpath('//input[@type="text"][@class="_17ePo copyable-text selectable-text"]')
        contact_seacrch_box.send_keys(i)
        user = driver.find_element_by_xpath("//span[@class='_3ko75 _5h6Y_ _3Whw5'][@title='{}']".format(i))
        user.click()
        sheet.write(row, 0, i)
        sheet.write(row, 1, "Succeed")
        sleep(1)
        row += 1
    except:
        sheet.write(row, 0, i)
        sheet.write(row, 1, "Failed")
        contact_seacrch_box.clear()
        sleep(1)
        row += 1

workbook.close()
driver.quit()

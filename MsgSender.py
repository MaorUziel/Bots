from selenium import webdriver
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from random import randint
import pandas as pd
from time import sleep
import xlsxwriter


options = FirefoxOptions()
driver = webdriver.Firefox(options=options, executable_path="C:\webdrivers\geckodriver.exe")
driver.get('https://www.messenger.com/')
email = input('Enter your email: ')
password = input('Enter your password: ')

email_id = driver.find_element_by_id('email')
email_id.send_keys(email)
password_id = driver.find_element_by_id('pass')
password_id.send_keys(password)
log = driver.find_element_by_id('loginbutton')
log.click()

df = pd.read_excel('pm.xlsx')
link_list = list(df['link'])
name_list = list(df['name'])

status_sheet = "SenderStatus.xlsx"
workbook = xlsxwriter.Workbook(status_sheet)
sheet = workbook.add_worksheet("Status")

sheet.write(0, 0, 'Name')
sheet.write(0, 1, 'Link')
sheet.write(0, 2, 'Status')


print("Enter your message after greeting: ")
print("When you're done, enter a single period on a line by itself.")
msg = []  # each line user input add as term to the msg list
while True:
    print("> ", end="")
    line = input()
    if line == ".":
        break
    msg.append(line)

i = 1
for link in link_list[0:5]:
    try:
        profile_id = (str(link.split('/')[-1:][0])).split('=')[-1:][0]
    except:
        try:
            profile_id = link.split('/')[-1:][0]
        except:
            sheet.write(i, 0, (name_list[(i - 1)]))
            sheet.write(i, 1, link)
            sheet.write(i, 2, 'Failed to create link')
            i += 1

    profile_link = 'https://www.messenger.com/t/' + profile_id
    driver.get(profile_link)

    try:
        msg_box = driver.find_element_by_xpath('//div[@aria-label="Type a message..."]')
        for line in msg:  # create in order to send multi line message
            ActionChains(driver).key_down(Keys.SHIFT).send_keys(Keys.ENTER).key_up(Keys.SHIFT).perform()
            msg_box.send_keys(line)
       # msg_box.send_keys(Keys.ENTER)
        sheet.write(i, 0, (name_list[(i-1)]))
        sheet.write(i, 1, link)
        sheet.write(i, 2, 'Succeed')
        sleep(randint(4, 16))
    except:
        sheet.write(i, 0, (name_list[(i-1)]))
        sheet.write(i, 1, link)
        sheet.write(i, 2, 'Failed')
        sleep(randint(4, 16))
    i += 1


driver.quit()
workbook.close()
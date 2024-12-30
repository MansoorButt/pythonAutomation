from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
import pandas as pd

driver_location = '/usr/bin/chromedriver'
binary_location = '/usr/bin/google-chrome'
option = webdriver.ChromeOptions()
option.binary_location = binary_location

driver = webdriver.Chrome(executable_path=driver_location, chrome_options=option)
driver.get('https://ti2ptih6lthnanlb.vercel.app/')

dataframe = pd.read_excel('data.xlsx')
entry = dataframe.iloc[0]

name_input = driver.find_element_by_name('fullname')
name_input.send_keys(entry['Full Name'])

passport_input = driver.find_element_by_name('passportNumber')
passport_input.send_keys(entry['PassportNumber'])

nationality_input = driver.find_element_by_name('nationality')
nationality_input.send_keys(entry['Nationality'])

email_input = driver.find_element_by_name('email')
email_input.send_keys(entry['Email'])


visa_xpath = "/html/body/div/div/form/div[1]/div[4]/button"
visaType_select = Select(driver.find_element_by_xpath(visa_xpath))
visaType_select.select_by_visible_text(entry['VisaType'])

time.sleep(3)
submit_btn = driver.find_element_by_xpath('/html/body/div/div/form/div[2]/button')
submit_btn.click()

time.sleep(2)
driver.quit()
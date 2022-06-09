# fill form with data from a csv file
from selenium import webdriver
import time, sys, csv

web = webdriver.Chrome()
web.get('https://www.cdicollege.ca/')

time.sleep(2)

# web.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input').send_keys('python')
# web.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
# time.sleep(2)

web.find_element_by_xpath('//*[@id="right-menu-section-in-header-id"]/div/a[1]').click()

time.sleep(2)

web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[1]/div/label[1]').click()

time.sleep(2)







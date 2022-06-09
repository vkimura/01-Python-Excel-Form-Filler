# fill form with data from a csv file
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

import time, sys, csv

web = webdriver.Chrome()
web.get('https://www.collegecdi.loc/')

time.sleep()

# web.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input').send_keys('python')
# web.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
# time.sleep(2)

web.find_element_by_xpath('//*[@id="right-menu-section-in-header-id"]/div/a').click() # click on the "Request Info" button

time.sleep(2)

web.find_element_by_xpath('//*[@id="int-yes2"]').click() # click on the "I am an international student" button

time.sleep(2)

web.find_element_by_xpath('//*[@id="res-no2"]').click() # click on the "I am not a resident of Canada" button

time.sleep(1)

web.find_element_by_xpath('//*[@id="CountryKey"]').send_keys('Canada') # enter "Canada" in the "Country" field

time.sleep(1)

web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[6]/input').send_keys('Timmy') # enter "Timmy" in the "First Name" field

#time.sleep(1)

web.find_elements_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[7]/input')[0].send_keys('Tom') # enter "Tom" in the "Last Name" field

#time.sleep(1)

web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[8]/input').send_keys('timmy@mailinator.com') # enter email

#time.sleep(1)

web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[9]/input').send_keys('123-343-3734') # enter phone number

#time.sleep(1)

web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[10]/input').send_keys('V5X 5RT') # enter postal code

time.sleep(1)

select = Select(web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[11]/select')) # define the "Program" drop-down

time.sleep(1)

select.select_by_visible_text('Gestion de l\'approvisionnement - LCA.FL') # select program

time.sleep(2)

web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/button').click() # click on the "Submit" button

time.sleep(10)

    #web.close()
    #web.quit()















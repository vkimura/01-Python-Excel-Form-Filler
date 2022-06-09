# fill form with data from a csv file
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from xlrd import open_workbook

wb = open_workbook("C:\\Users\\vkimura\\Documents\\Projects\\2022\\05\\09\\2022-06-08-1-colon-20PM.xls")
sheet = wb.sheet_by_index(0)
# sheet.cell_value(0, i)

import time, sys, csv

#get column names into a list
sheet.cell_value(0, 0)
columns = []
for i in range(sheet.ncols):
    columns.append(sheet.cell_value(0, i))

#get index from array of column names
# def get_index(column_name):
#     for i in range(sheet.ncols):
#         if sheet.cell_value(0, i) == column_name:
#              return i
indexSite = columns.index("Site")
indexURL = columns.index("URL")
indexIntlStudent = columns.index("Int'l student?")
indexStudyPermit = columns.index("Study permit")
indexRefugeeStatus = columns.index("Refugee status")
indexResideInCanada = columns.index("Reside in Canada")
indexCountry = columns.index("Country")
indexFirstName = columns.index("First Name")
indexLastName = columns.index("Last Name")
indexEmail = columns.index("Email")
indexPhone = columns.index("Phone")
indexPostal = columns.index("Postal")
indexProgram = columns.index("Program")
indexLandingPage = columns.index("Landing Page")
indexInLeadsTable = columns.index("In leads Table")
indexMyCollegeLeads = columns.index("MyCollegeLeads.ca")

print(indexURL)

web = webdriver.Chrome()
# column 2 in Excel
#web.get('https://www.collegecdi.loc/')
web.get(sheet.cell_value(0, 2))

time.sleep()

# web.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input').send_keys('python')
# web.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
# time.sleep(2)

web.find_element_by_xpath('//*[@id="right-menu-section-in-header-id"]/div/a').click() # click on the "Request Info" button

time.sleep(2)

# click on the "I am an international student" button - column 5 in Excel
if (indexSite == "cdicollege"):
    if (sheet.cell_value(0, indexIntlStudent) == "Yes"):
        web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[1]/div/label[1]').click()
    else:
        web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[1]/div/label[2]').click()
elif (indexSite == "collegecid"):
    if (sheet.cell_value(0, indexIntlStudent) == "Yes"):
        web.find_element_by_xpath('//*[@id="int-yes2"]').click()
    else:
        web.find_element_by_xpath('//*[@id="int-no2"]').click()

# //*[@id="submitRequestInfo"]/div[2]/div[3]/div[1]/div/label[1] - yes (CDICollege XPath)
# //*[@id="submitRequestInfo"]/div[2]/div[3]/div[1]/div/label[2] - no (CDICollege XPath)
#web.find_element_by_xpath('//*[@id="int-yes2"]').click()
#//*[@id="int-no2"] - no  (CollegeCDI XPath)

#click on the "Do you have a study permit in Canada?" button - column 6 in Excel
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[2]/div/label[1] - yes (CDICollege XPath)
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[2]/div/label[2] - no (CDICollege XPath)

#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[2]/div/label[1] - yes (CollegeCDI XPath)
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[2]/div/label[2] - no (CollegeCDI XPath)

#click on the "Do you have a refugee status in Canada?" button - column 7 in Excel
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[3]/div/label[1] - yes (CDICollege XPath)
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[3]/div/label[2] - no (CDICollege XPath)

#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[3]/div/label[1] - yes (CollegeCDI XPath)
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[3]/div/label[2] - no (CollegeCDI XPath)

time.sleep(2)

#click on "Do you have a Canadian address?" button - column 8 in Excel
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[4]/div/label[1] - yes (CDICollege XPath)
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[4]/div/label[2] - no (CDICollege XPath)

#//*[@id="res-yes2"] - yes (CollegeCDI XPath)
web.find_element_by_xpath('//*[@id="res-no2"]').click() # click on the "I am not a resident of Canada" button

time.sleep(1)

# enter "Canada" in the "Country" field - Country drop down - column 9 in Excel
# web.find_element_by_xpath('//*[@id="CountryKey"]').send_keys('Canada')
web.find_element_by_xpath('//*[@id="CountryKey"]').send_keys(sheet.cell_value(0, indexCountry))

time.sleep(1)

# enter "Timmy" in the "First Name" field - column 10 in Excel
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[6]/input - first name (CollegeCDI XPath)
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[6]/input - first name (CDICollege XPath)
if (indexSite == "cdicollege" || indexSite == "collegecdi"):
    # web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[6]/input').send_keys('Timmy')
    web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[6]/input').send_keys(sheet.cell_value(0, indexFirstName))

#time.sleep(1)

# enter "Tom" in the "Last Name" field - column 11 in Excel
web.find_elements_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[7]/input')[0].send_keys('Tom') 

#time.sleep(1)

# enter email in the "Email" field - column 12 in Excel
web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[8]/input').send_keys('timmy@mailinator.com')

#time.sleep(1)

# enter phone number in the "Phone" field - column 13 in Excel
web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[9]/input').send_keys('123-343-3734')

#time.sleep(1)

# enter postal in the "Postal Code" field - column 14 in Excel
web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[10]/input').send_keys('V5X 5RT')

time.sleep(1)

# define the "Program" drop-down
select = Select(web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[11]/select'))

time.sleep(1)

# select program from the drop-down - column 15 in Excel
select.select_by_visible_text('Gestion de l\'approvisionnement - LCA.FL')

time.sleep(2)

web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/button').click() # click on the "Submit" button

time.sleep(10)

    #web.close()
    #web.quit()















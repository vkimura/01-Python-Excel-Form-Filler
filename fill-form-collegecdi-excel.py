# fill form with data from a csv file
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from xlrd import open_workbook

wb = open_workbook("C:\\Users\\vkimura\\Documents\\Projects\\2022\\05\\09\\2022-06-08-1-colon-20PM.xls")
sheet = wb.sheet_by_index(0) # sheet 0
# sheet.cell_value(0, i)

import time, sys, csv

#get column names into a list
sheet.cell_value(0, 0) #get first row (i.e. column names)
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
indexForm = columns.index("Form")
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
indexInLeadsTable = columns.index("In Leads Table")
indexMyCollegeLeads = columns.index("MyCollegeLeads.ca")
setRowNumber = 10    # start at row 5
setSleepTimeLong = 2
setSleepTime = 1

print(indexURL)

web = webdriver.Chrome()
# column 2 in Excel
#web.get('https://www.collegecdi.loc/')
# web.get(sheet.cell_value(0, 2))
url = sheet.cell_value(setRowNumber, indexURL)
web.get(sheet.cell_value(setRowNumber, indexURL))

time.sleep(setSleepTimeLong)

# web.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input').send_keys('python')
# web.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
# time.sleep(2)

#//*[@id="header-section-v2-id"]/div[1]/div/div/div[2]/a[1] - click on the "Apply Now" button (CDICollege)
#//*[@id="header-section-v2-id"]/div[1]/div/div/div[2]/a[1]  - click on the "Inscrivez-vous" button (CollegeCDI)
formName = sheet.cell_value(setRowNumber, indexForm)
if (formName == "Demande d'info" or formName =="Request Information"):
    web.find_element_by_xpath('//*[@id="right-menu-section-in-header-id"]/div/a').click() # click on the "Request Info"/"Demande d'info" button
elif (formName == "Inscrivez-vous" or formName =="Apply Now"):
    web.find_element_by_xpath('//*[@id="header-section-v2-id"]/div[1]/div/div/div[2]/a[1]').click() # click on the "Apply Now"/"Inscrivez-vous" button

time.sleep(setSleepTimeLong)

# click on the "I am an international student" button - column 5 in Excel
siteName = sheet.cell_value(setRowNumber, indexSite)
intlStudentVal = sheet.cell_value(setRowNumber, indexIntlStudent)
if (siteName == "cdicollege"):
    if (intlStudentVal == "yes"):
        web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[1]/div/label[1]').click()
    else:
        web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[1]/div/label[2]').click()
elif (siteName == "collegecdi"):
    if (intlStudentVal.lower() == "yes"):
        web.find_element_by_xpath('//*[@id="int-yes2"]').click()
    else:
        web.find_element_by_xpath('//*[@id="int-no2"]').click()

# //*[@id="submitRequestInfo"]/div[2]/div[3]/div[1]/div/label[1] - yes (CDICollege XPath)
# //*[@id="submitRequestInfo"]/div[2]/div[3]/div[1]/div/label[2] - no (CDICollege XPath)
#web.find_element_by_xpath('//*[@id="int-yes2"]').click()
#//*[@id="int-no2"] - no  (CollegeCDI XPath)



time.sleep(setSleepTime)    

#if the student is International, then the following other options apply; otherise, the other fields are not displayed.
if (intlStudentVal == "yes"):
        
    #click on the "Do you have a study permit in Canada?" button - column 6 in Excel
        #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[2]/div/label[1] - yes (CDICollege XPath)
        #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[2]/div/label[2] - no (CDICollege XPath)

        #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[2]/div/label[1] - yes (CollegeCDI XPath)
        #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[2]/div/label[2] - no (CollegeCDI XPath)
    studyPermitVal = sheet.cell_value(setRowNumber, indexStudyPermit)
    if (studyPermitVal.lower() == "yes"):
        web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[2]/div/label[1]').click()
    elif (studyPermitVal.lower() == "no"):
        web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[2]/div/label[2]').click()

    #click on the "Do you have a refugee status in Canada?" button - column 7 in Excel
        #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[3]/div/label[1] - yes (CDICollege XPath)
        #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[3]/div/label[2] - no (CDICollege XPath)

        #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[3]/div/label[1] - yes (CollegeCDI XPath)
        #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[3]/div/label[2] - no (CollegeCDI XPath)

    time.sleep(setSleepTime)    

    refugeeStatusVal = sheet.cell_value(setRowNumber, indexRefugeeStatus)
    if (refugeeStatusVal.lower() == "yes"):
        web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[3]/div/label[1]').click()
    elif (refugeeStatusVal.lower() == "no"):
        web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[3]/div/label[2]').click()

    time.sleep(setSleepTime)

    #click on "Do you have a Canadian address?" button - column 8 in Excel
        #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[4]/div/label[1] - yes (CDICollege XPath)
        #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[4]/div/label[2] - no (CDICollege XPath)

        #//*[@id="res-yes2"] - yes (CollegeCDI XPath)
        #web.find_element_by_xpath('//*[@id="res-no2"]').click() # click on the "I am not a resident of Canada" button

    if (siteName == "cdicollege"):
        if (sheet.cell_value(setRowNumber, indexResideInCanada).lower() == "yes"):
            web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[4]/div/label[1]').click() # click on the "yes" button
        else:
            web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[4]/div/label[2]').click() # click on the "no" button
    elif (siteName == "collegecdi"):
        if (sheet.cell_value(setRowNumber, indexResideInCanada).lower() == "yes"):
            web.find_element_by_xpath('//*[@id="res-yes2"]').click() # click on the "yes" button
        else:
            web.find_element_by_xpath('//*[@id="res-no2"]').click() # click on the "no" button  

    time.sleep(1)

    #check if country list selection is displayed
    strStyleCountryList = web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[5]').get_attribute('style')
    if (strStyleCountryList.find('display: none') == -1):
        selectCountryList = Select(web.find_element_by_xpath('//*[@id="CountryKey"]')) #can select by name using this variable
        web.find_element_by_xpath('//*[@id="CountryKey"]').send_keys(sheet.cell_value(setRowNumber, indexCountry))

    # enter "Canada" in the "Country" field - Country drop down - column 9 in Excel
    # web.find_element_by_xpath('//*[@id="CountryKey"]').send_keys('Canada')
        # web.find_element_by_xpath('//*[@id="CountryKey"]').send_keys(sheet.cell_value(setRowNumber, indexCountry))

time.sleep(1)

# enter "Timmy" in the "First Name" field - column 10 in Excel
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[6]/input - first name (CollegeCDI XPath)
#//*[@id="submitRequestInfo"]/div[2]/div[3]/div[6]/input - first name (CDICollege XPath)
if (siteName == "cdicollege" or siteName == "collegecdi"):
    # web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[6]/input').send_keys('Timmy')
    web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[6]/input').send_keys(sheet.cell_value(setRowNumber, indexFirstName))

#time.sleep(1)

# enter "Tom" in the "Last Name" field - column 11 in Excel
if (siteName == "cdicollege" or siteName == "collegecdi"):
    # web.find_elements_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[7]/input')[0].send_keys('Tom') 
    web.find_elements_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[7]/input')[0].send_keys(sheet.cell_value(setRowNumber, indexLastName)) 

#time.sleep(1)

# enter email in the "Email" field - column 12 in Excel
if (siteName == "cdicollege" or siteName == "collegecdi"):
    # web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[8]/input').send_keys('timmy@mailinator.com')
    web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[8]/input').send_keys(sheet.cell_value(setRowNumber, indexEmail))

#time.sleep(1)

# enter phone number in the "Phone" field - column 13 in Excel
if (siteName == "cdicollege" or siteName == "collegecdi"):
    # web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[9]/input').send_keys('123-343-3734')
    web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[9]/input').send_keys(sheet.cell_value(setRowNumber, indexPhone))

#time.sleep(1)

# enter postal in the "Postal Code" field - column 14 in Excel
if (siteName == "cdicollege" or siteName == "collegecdi"):
    # web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[10]/input').send_keys('V5X 5RT')
    web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[10]/input').send_keys(sheet.cell_value(setRowNumber, indexPostal))

time.sleep(1)

# define the "Program" drop-down
    #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[11]/select - program (CollegeCDI XPath)
    #//*[@id="submitRequestInfo"]/div[2]/div[3]/div[11]/select - program (CDICollege XPath)
    # strStyleCountryList = web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[5]').get_attribute('style')
    # select = Select(web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[11]/select'))
    # web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[11]/div/select').click()

time.sleep(1)

# select program from the drop-down - column 15 in Excel
selectProgram = Select(web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/div[3]/div[11]/select'))
selectProgram.select_by_visible_text(sheet.cell_value(setRowNumber, indexProgram))
# try:
#     select
# except NameError:
#     var_exists = False
# else:
#     var_exists = True
#     # select.select_by_visible_text('Gestion de l\'approvisionnement - LCA.FL')
#     select.select_by_visible_text(sheet.cell_value(setRowNumber, indexProgram))

time.sleep(2)

#web.find_element_by_xpath('//*[@id="submitRequestInfo"]/div[2]/button').click() # click on the "Submit" button

time.sleep(10)

    #web.close()
    #web.quit()















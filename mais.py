#Import libraries selenium and pandas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support.ui import Select
import pandas as pd
#import seaborn as sns
import matplotlib.pyplot as plt
from pandas import ExcelWriter

driver = webdriver.Firefox()
driver.get('url_name.com') #Open mais page
driver.implicitly_wait(30)

user_name = driver.find_element_by_name('username') #Find the user name tag element
user_name.clear() #Clear the initial values
user_name.send_keys('UserID')#Change USER ID
driver.implicitly_wait(30)

password= driver.find_element_by_name('password')#Find the password tag element
password.clear()
password.send_keys('Password')#Change USER PASSWORD
driver.implicitly_wait(30)

login = driver.find_element_by_id('login-submit')#Find the login-submit button id and click
login.click()


##########################################
frame = driver.find_elements_by_tag_name('frame')
#print frame[:]
driver.switch_to_frame(frame[1])

frame_left = driver.find_elements_by_tag_name('frame')
#print frame_left[:]
driver.switch_to_frame(frame_left[1])

frame_required = driver.find_elements_by_tag_name('frame')
#print frame_required[:]
driver.switch_to_frame(frame_required[0])
driver.find_element_by_link_text('Job Request Form').click()

##########################################

driver.switch_to_default_content()
frame = driver.find_elements_by_tag_name('frame')
driver.switch_to_frame(frame[1])
frame_left = driver.find_elements_by_tag_name('frame')
driver.switch_to_frame(frame_left[1])
frame_required = driver.find_elements_by_tag_name('frame')
driver.switch_to_frame(frame_required[1])


name_assy = driver.find_element_by_id('jobName')
name_assy.send_keys('Name')

df = pd.read_csv('mais_trial.csv') #CSV NOT EXCEL FILE , STORE IN SAME FOLDER
no = df.shape[0]                   #ROWS OF CSV FILE EXCLUDING HEADER ROW
no_of_drwgs = driver.find_element_by_id('drawingSets')
no_of_drwgs.clear()
no_of_drwgs.send_keys(str(no))


driver.find_element_by_xpath("//select[@name='project']/option[text()='Project_Name']").click()#Change PROJECT
driver.find_element_by_xpath("//select[@name='model']/option[text()='Model No']").click()#Change MODEL
driver.find_element_by_xpath("//select[@name='inspectionRequired']/option[text()='Yes']").click()#Change INSOECTION REQUIRED/NOT
driver.find_element_by_name('dateRequiredInput').send_keys('28-02-2018')#Change DATE REQURIED (NOT WORKING)
driver.find_element_by_xpath("//select[@name='aauthority']/option[text()='Name of person']").click()#Change APPROVING AUTHORITY


#################################################################################################
for i in range(no):
	driver.find_element_by_id('description_'+str(i+1)).send_keys(df['Description'][i])
	driver.find_element_by_id('drawingNo_'+str(i+1)).send_keys(df['Dwg No.'][i])
	driver.find_element_by_id('revision_'+str(i+1)).clear()
	driver.find_element_by_id('revision_'+str(i+1)).send_keys(str(df['Revision'][i]))
	driver.find_element_by_id('quantity_'+str(i+1)).clear()
	driver.find_element_by_id('quantity_'+str(i+1)).send_keys(str(df['Quantity'][i]))
	
	if df['Raw Material'][i][0].upper() in ['A','S','T','D','V']:
		text_1 = df['Raw Material'][i].replace(' ','').upper()
		df['Raw Material'][i] = text_1	
	elif df['Raw Material'][i][0].upper() in ['B']:
		text_1 = 'BERYLIUM COPPER'
		df['Raw Material'][i] = text_1	
	if df['Shape'][i][0].upper() in ['R','C']: 
		text_2 = df['Shape'][i][0].upper() + df['Shape'][i][1:].lower()
		df['Shape'][i] = text_2
	elif df['Shape'][i][0:2].upper() in ['SH']:
		text_2 =  df['Shape'][i][:].upper() + str('/PLATE')
		df['Shape'][i] = text_2		
	elif df['Shape'][i][0:2].upper() in ['SQ']:
		text_2 = df['Shape'][i][0].upper() + df['Shape'][i][1:].lower() 	 
		df['Shape'][i] = text_2
	
	rm = Select(driver.find_element_by_id('rawMaterial_'+str(i+1)))
	rm.select_by_visible_text(str(text_1))
	sh = Select(driver.find_element_by_id('shape'))
	sh.select_by_visible_text(str(text_2))
	
	driver.find_element_by_name('size1').clear()
	driver.find_element_by_name('size1').send_keys(str(df['size1'][i]))
	driver.find_element_by_name('size2').clear()
	driver.find_element_by_name('size2').send_keys(str(df['size2'][i]))
	
	driver.find_element_by_name('size3').send_keys(str(df['size3'][i]))
	driver.find_element_by_name('mtrlQty').clear()
	driver.find_element_by_name('mtrlQty').send_keys(str(df['Mquantity'][i]))
	driver.find_element_by_name('purOrderNo').clear()
	driver.find_element_by_name('purOrderNo').send_keys(str(df['PO'][i]))
	driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/button[1]").click()
	



# MAIS Automation script
# Author: Kartik Sah
# Year: 2018

#Import libraries selenium and pandas
try:
	from selenium import webdriver
	from selenium.webdriver.common.by import By
	from selenium.webdriver.support import expected_conditions as EC
	from selenium.webdriver.support.ui import WebDriverWait as wait
	from selenium.webdriver.support.ui import Select
	import pandas as pd
	from pandas import ExcelWriter
	import getpass
	import datetime
except ImportError:
	print('Module not Found')

'''
Name of the Approving Authority, and the corresponding number.
'''	
approver = {1 : 'KESHAVA MURTHY K A', 2 : 'RAVINDRAN S', 3 : 'SHANKARA A', 4 : 'SURESHA KUMAR H N'}
folder_add = "C:\Python27\Drawing_pdf\TEMP"+'\\'

############################################################################################################################################
def decorator_func(use_func):
	import logging
	import time
	logging.basicConfig(filename='{}.log'.format('Mais_Log_Info'), level = logging.INFO)
	
	if use_func.__name__ == 'open_start_page':
		logging.info('************** New Session **************')
		
	def wrapper_func(*args,**kwargs):
		time_start = time.time()
		result = use_func(*args)
		time_end = time.time()
		delta_time = time_end - time_start
		logging.info(str(use_func.__name__) + ':{} , Time taken to complete: {:.2f} seconds'.format(args[1],delta_time))
		return result
	return wrapper_func


############################################################################################################################################
def password_check(password):
	condition = raw_input('Show Password (Y/N):')
	if condition == 'Y':
		print(password)

#############################################################################################################################################
# check_func = decorator_func(check_func) -- is equivalent to @decorator_func
@decorator_func
def open_start_page(driver,user_name,password,url):
	
	driver.get(url) #Open mais page
	driver.implicitly_wait(30)

	user_name_tab= driver.find_element_by_id('text1') #Find the user name tag element
	user_name_tab.clear() #Clear the initial values
	user_name_tab.send_keys(user_name) #Change USER ID
	driver.implicitly_wait(30)

	password_tab= driver.find_element_by_id('text2')#Find the password tag element
	password_tab.clear()
	password_tab.send_keys(str(password))#Change USER PASSWORD
	driver.implicitly_wait(30)

	login = driver.find_element_by_id('login-submit')#Find the login-submit button id and click
	login.click()

	jobrequestform(driver)
	return driver

##############################################################################################################################################
def jobrequestform(driver):
	frame = driver.find_elements_by_tag_name('frame')
	driver.switch_to_frame(frame[1])
	frame_left = driver.find_elements_by_tag_name('frame')
	driver.switch_to_frame(frame_left[1])
	frame_required = driver.find_elements_by_tag_name('frame')
	driver.switch_to_frame(frame_required[0])
	driver.find_element_by_link_text('Job Request Form').click()


###############################################################################################################################################
@decorator_func
def first_page_data_table(driver,total_drawings,df):
	for i in range(total_drawings):
		driver.find_element_by_id('description_'+str(i+1)).send_keys(df['Description'][i])
		driver.find_element_by_id('drawingNo_'+str(i+1)).send_keys(df['Dwg No.'][i])
		driver.find_element_by_id('revision_'+str(i+1)).clear()
		driver.find_element_by_id('revision_'+str(i+1)).send_keys(str(df['Revision'][i]))
		driver.find_element_by_id('quantity_'+str(i+1)).clear()
		driver.find_element_by_id('quantity_'+str(i+1)).send_keys(str(df['Quantity'][i]))
		text_1 = df['Raw Material'][i][:]

		text_2 = df['Shape'][i][:]
		wait(driver,1000.0)
		
		rm = Select(driver.find_element_by_id('rawMaterial_'+str(i+1)))
		rm.select_by_visible_text(str(text_1))
		sh = Select(driver.find_element_by_id('shape'))
		sh.select_by_visible_text(str(text_2))

		wait(driver,1000.0)
		
		driver.find_element_by_name('size1').send_keys(str(df['size1'][i]))
		driver.find_element_by_name('size2').send_keys(str(df['size2'][i]))
		driver.find_element_by_name('size3').send_keys(str(df['size3'][i]))
		driver.find_element_by_name('mtrlQty').send_keys(str(df['Mquantity'][i]))
		driver.find_element_by_name('purOrderNo').clear()
		driver.find_element_by_name('purOrderNo').send_keys(str(df['PO'][i]))
		driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/button[1]").click()
		driver.implicitly_wait(30)
	return drawing_upload_page(driver,df)

#############################################################################################################
def first_data_page(driver,assy_name,file_name,approver,aaprover):
	driver.switch_to_default_content()
	frame = driver.find_elements_by_tag_name('frame')
	driver.switch_to_frame(frame[1])
	frame_left = driver.find_elements_by_tag_name('frame')
	driver.switch_to_frame(frame_left[1])
	frame_required = driver.find_elements_by_tag_name('frame')
	driver.switch_to_frame(frame_required[1])


	name_assy = driver.find_element_by_id('jobName')
	name_assy.send_keys(assy_name)

	df = pd.read_excel('C:/Python27/Upload_excel/' + file_name + '.xlsx', index= False) #EXCEL FILE , STORE IN SAME FOLDER
	current_time = datetime.datetime.now()
	new_file_name = file_name + '_current_time.day'+ '/current_time.month'+ '/current_time.year'+ '_current_time.hour'+ '_current_time.minute'+ '_current_time.second'

	total_drawings = df.shape[0]                  #ROWS OF EXCEL FILE EXCLUDING HEADER ROW
	no_of_drwgs = driver.find_element_by_id('drawingSets')
	no_of_drwgs.clear()
	no_of_drwgs.send_keys(str(total_drawings))

	driver.find_element_by_xpath("//select[@name='project']/option[text()='--Please select --']").click()
	#driver.find_element_by_xpath("//select[@name='project']/option[text()='--Please select --']").click()#Change PROJECT
	driver.find_element_by_xpath("//select[@name='model']/option[text()='FLIGHT']").click()#Change MODEL
	driver.find_element_by_xpath("//select[@name='inspectionRequired']/option[text()='Yes']").click()#Change INSPECTION REQUIRED/NOT
	driver.find_element_by_name('dateRequiredInput').click()#Change DATE REQURIED (NOT WORKING)
	driver.find_element_by_xpath('/html/body/div[3]/div/div/select[1]').click()
	driver.find_element_by_xpath('/html/body/div[3]/div/div/select[1]/option[4]').click()
	driver.find_element_by_xpath('/html/body/div[3]/table/tbody/tr[3]/td[3]/a').click()
	driver.find_element_by_xpath("//select[@name='aauthority']/option[text()=" + "\'"+ str(approver[int(aaprover)]) + "\'"  "]").click()#Change APPROVING AUTHORITY
	first_page_data_table(driver,total_drawings,df)
	
###############################################################################################################################################

def drawing_upload_page(driver,df):
	if 'Y' == raw_input('Change the  Date and press (Y):').upper():
		driver.find_element_by_id('indentEntry').click()
		driver.switch_to_default_content()
		frame = driver.find_elements_by_tag_name('frame')
		driver.switch_to_frame(frame[1])
		frame_left = driver.find_elements_by_tag_name('frame')
		driver.switch_to_frame(frame_left[1])
		frame_required = driver.find_elements_by_tag_name('frame')
		driver.switch_to_frame(frame_required[1])

		upload = driver.find_elements_by_id('uploadForm_drawingFile')
		for i in range(len(upload)):
			wait(driver,10.0)
			upload[i].send_keys(folder_add + str(df['Filename'][i]) + '.pdf')

def password_enter():
	while True:
		password = getpass.getpass('Enter the Password:')
		password_check(password)
		condition = raw_input('Re Enter the password(Y/N):')
		if condition.upper() == 'N':
			break
	return password


if __name__ == '__main__':
	
	url = 'http://10.21.6.100:9090/mais/index.action'

	user_name = raw_input('Enter User Name:')
	password = password_enter()
	assy_name = raw_input('Enter the Assembly Name:')
	file_name = raw_input('Enter the filename:')
	print(approver)
	aaprover  = int(raw_input('Choose the approver (Select the Number corresponding to the Name): '))
	if aaprover not in [1,2,3,4]:
		aaprover = 3

	driver = webdriver.Firefox()
	open_start_page(driver,user_name,password,url)
	first_data_page(driver,assy_name,file_name,approver,aaprover)
	wait(driver,15)
	

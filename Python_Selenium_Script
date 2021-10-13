from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pynput.keyboard import Key, Controller
import time
import pygetwindow as gw
import openpyxl
import pyautogui

path = "C:\\Users\MEMaldonado\OneDrive - United Biologics LLC\Shared Order Entry Files\Label_Printing.xlsx"

keyboard = Controller()

def press_release_char(char):
	keyboard.press(char)
	keyboard.release(char)		


driver = webdriver.Chrome('./chromedriver')
driver.maximize_window()
driver.get("https://prod.rxdispense.com/#/user/home/index")
#wait for user to login
try:
    	element = WebDriverWait(driver, 30).until(EC.title_is("RxDispense Home"))
except:
	print("Could not locate home screen. Please ensure you log in.")
driver.find_element_by_partial_link_text('Rx - New').click()

while True:
	#open excel sheet
	wb_obj = openpyxl.load_workbook(path) 
	sheet_obj = wb_obj.active 

	#find next unprinted label and print
	row = sheet_obj.max_row
	column = sheet_obj.max_column
	for i in range(2, row+1):
		if sheet_obj.cell(row = i, column = 2).value is None:
			rxNumber = sheet_obj.cell(row = i, column = 1).value
			#convert Rx number to string
			stringRxNumber = str(rxNumber)
			clickableSearchBox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'searchTextBox')))
			driver.find_element_by_id('searchTextBox').click()

			for element in range(0, len(stringRxNumber)):
				press_release_char(stringRxNumber[element])

			#click broad search button
			driver.find_element_by_xpath('//*[@id="dashboardMainContent"]/div[2]/div/div/div/div[2]/div[1]/div').click()

			#wait until result loads and select
			WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="dashboardGrid"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[3]'), stringRxNumber))
			driver.find_element_by_xpath('//*[@id="dashboardGrid"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[3]').click()
	
			#wait for Rx result to load
			WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'mainRxSidePanel')))

			##PRINT RX LABEL

			#select Choose Label
			driver.find_element_by_xpath('//*[@id="mainRxSidePanel"]/div/div[2]').click()

			#wait for print screen to load
			WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mainBody"]/div[5]/div/div[2]/div/div/div/div[6]/div/div/div[1]/div/table/tbody/tr[7]/td[1]')))

			#select AlleReach Rx Label
			driver.find_element_by_xpath('//*[@id="mainBody"]/div[5]/div/div[2]/div/div/div/div[6]/div/div/div[1]/div/table/tbody/tr[7]/td[1]').click()
	
			#Wait for print screen to load and press Print
			#clickableSearchBox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/print-preview-app')))
			#driver.find_element_by_xpath('//*[@id="sidebar"]//print-preview-button-strip//div/cr-button[1]').click()
	
			time.sleep(12)
			#get to 'Destination'
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.enter)
			#Select 'See more'
			press_release_char(Key.down)
			press_release_char(Key.down)
			press_release_char(Key.down)
			press_release_char(Key.down)
			press_release_char(Key.enter)
			#Type in search bar and pick Rx Label 3
			time.sleep(4)
			keyboard.type('Rx Label 3')
			time.sleep(2)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.enter)
			#select 'More settings'
			time.sleep(2)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.enter)
			#Select 'Scale' and set to 'Fit to printable area'
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.up)
			press_release_char(Key.up)
			press_release_char(Key.down)
			press_release_char(Key.enter)
			press_release_char(Key.enter)
			#Select 'Print' and press Enter
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			time.sleep(2)
			press_release_char(Key.enter)
	
			cellObjectVialLabel = sheet_obj.cell(row = i, column = 2)
			cellObjectVialLabel.value = "Label printed"
			wb_obj.save(path) 
		
			##PRINT VIAL LABEL

			#select Choose Label
			driver.find_element_by_xpath('//*[@id="mainRxSidePanel"]/div/div[2]').click()
	
			#wait for print screen to load
			clickableSearchBox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mainBody"]/div[5]/div/div[2]/div/div/div/div[6]/div/div/div[1]/div/table/tbody/tr[7]/td[1]')))

			#select AlleReach Vial Label
			driver.find_element_by_xpath('//*[@id="mainBody"]/div[5]/div/div[2]/div/div/div/div[6]/div/div/div[1]/div/table/tbody/tr[8]').click()

			time.sleep(10)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.enter)
			#Select 'See more'
			press_release_char(Key.down)
			press_release_char(Key.down)	
			press_release_char(Key.down)
			press_release_char(Key.down)
			press_release_char(Key.enter)
			#Type in search bar and pick Rx Label 1
			time.sleep(2)
			keyboard.type('large vial label 1')
			time.sleep(2)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.enter)
			#Select 'Scale' and set to 'Fit to printable area'
			time.sleep(2)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			press_release_char(Key.up)
			press_release_char(Key.up)
			press_release_char(Key.down)
			press_release_char(Key.enter)
			press_release_char(Key.enter)
			#Select 'Print' and press Enter
			press_release_char(Key.tab)
			press_release_char(Key.tab)
			time.sleep(2)
			press_release_char(Key.enter)

			cellObjectVialLabel = sheet_obj.cell(row = i, column = 3)
			cellObjectVialLabel.value = "Label printed"
			wb_obj.save(path)
	
			#Close Rx File
			driver.find_element_by_xpath('//*[@id="closeButton"]/div').click()
		
	#close excel sheet
	wb_obj.close()
	clickableSearchBox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'searchTextBox')))
	driver.find_element_by_id('searchTextBox').click()
	press_release_char(Key.enter)
	time.sleep(120)

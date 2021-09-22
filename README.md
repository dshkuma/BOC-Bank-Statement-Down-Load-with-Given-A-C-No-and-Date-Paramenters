from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from bs4 import BeautifulSoup
import math
from selenium.webdriver.support import expected_conditions as expect
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook
from openpyxl import load_workbook

print(' [2]  : ACC1')
print(' [3] : ACC2')

optacc = input("Enter Bank Account Option : ")

sdat = input("Enter start Date 'dd-mm-yyyy' : ")
edat = input("Enter End Date 'dd-mm-yyyy' : ")

cont = input("Do you want to Continue : 1 ")

print(sdat)
print(edat)

switcher = {
    2: 'ACC1',
    3: 'ACC2',
 
}


ACCNO = switcher.get(int(optacc), "nothing")
print(ACCNO)
if cont == str(1):
   
    service = Service('E:/PythonD/NMKPY/chromedriver_win32/chromedriver.exe')  # Optional argument, if not specified will search path.
    service.start()
    driver = webdriver.Remote(service.service_url)
    
    df = ""
 
    driver.get('https://online.boc.lk/T001/channel.jsp')
    
    action = ActionChains(driver);

    US_ID = driver.find_element_by_id('fldLoginUserId')
    US_ID.send_keys('UID')
    
    PWD = driver.find_element_by_id('SKBPassword')
    PWD.send_keys('PWD')
    window_before = driver.window_handles[0]
    
 
    login_button = driver.find_element_by_xpath('//*[@id="section"]/div[5]/div/div/a[2]/input')
    

    login_button.click()
    #driver.close()
    window_after = driver.window_handles[1]
    
    driver.switch_to.window(window_after)
    
    (By.XPATH, "//*[@id='cmdLogout']")))
    
   
    
    lstAcc=WebDriverWait(driver, 120, 1).until(
            expect.visibility_of_element_located(
            (By.XPATH, '//*[@id="L1scrollWrapper"]/ul/li[3]/div[1]')))

    action.move_to_element(lstAcc).perform()

    
    lstAcc1=WebDriverWait(driver, 120, 1).until(
            expect.visibility_of_element_located(
            (By.XPATH, '//*[@id="L1scrollWrapper"]/ul/li[3]/div[3]/div/div/li[2]/ul/li[3]')))
    
    lstAcc1.click()
     
 
    
    driver.switch_to.frame(driver.find_element_by_id("frame_AAT1"))
    
  
    
    
    
    selacc=WebDriverWait(driver, 120, 1).until(
            expect.visibility_of_element_located(
            (By.XPATH, '//*[@id="Account Number #[^ ]"]')))
 
    actions = ActionChains(driver)
    actions.move_to_element(selacc)
    actions.click(selacc)
    actions.perform()
 
    
    acno=WebDriverWait(driver, 120, 1).until(
            expect.visibility_of_element_located(
            (By.XPATH,  "//*[@id='Account Number #[^ ]']/option["+optacc+"]")))

    acno.click()
   
    
    stdat=WebDriverWait(driver, 120, 1).until(
            expect.visibility_of_element_located(
            (By.XPATH, '//*[@id="fldfromdate"]')))
 
    stdat.click()
  
    driver.execute_script("arguments[0].value = " + str(sdat)+";",  stdat)
    
    endat=WebDriverWait(driver, 120, 1).until(
            expect.visibility_of_element_located(
            (By.XPATH, '//*[@id="fldtodate"]')))
 
    endat.click()
 
    driver.execute_script("arguments[0].value = " + str(edat)+";",  endat)
    
    btnSubm=WebDriverWait(driver, 120, 1).until(
            expect.visibility_of_element_located(
            (By.XPATH, '//*[@id="contentarea"]/div/div/ul/li/input')))
    
    
  
    btnSubm.click()
    
    lno = 0
    elm = driver.find_element_by_css_selector("#maintable > tbody > tr:nth-child(2) > td > div.buttonarea > ul > li > input")
    while elm:	 
     html=driver.page_source
     soup=BeautifulSoup(html,'html.parser')
     div=soup.select_one("#maintable") # CSS Selector
     table=pd.read_html(str(div))
     frames = [table[0]]
     result=pd.concat(frames,ignore_index=True)
     df = pd.DataFrame(result)
     #print(df)
     # print(x)
     if lno==0:
      BOC_ST=pd.DataFrame(df.iloc[4:])
    		# print(Sam_MF)
     else:
      BOC_ST= BOC_ST.append(df.iloc[5:], ignore_index=True)
      sleep(2)
     lno = lno + 1
     try:
      elm = driver.find_element_by_css_selector("#maintable > tbody > tr:nth-child(2) > td > div.buttonarea > ul > li > input")
     except:
      print('Button End')
      break   
     if  elm:
      elm.click()
     else:
      print('End of Button')
     #elm.send_keys(Keys.COMMAND, Keys.ENTER, 'H')
    print('Pages :' + str(lno))
    
    wb = Workbook()
    ws =  wb.active
    ws.title = ACCNO
    flnam = 'D:/Temp/Bank Statem BOC '+ ACCNO + ' From ' + sdat + 'To ' + edat +'.xlsx'
    print(flnam)
    wb.save(filename = flnam)
    BOC_ST.to_excel(flnam)
    #driver.quit()
else:
    
    print('User Cancelled')

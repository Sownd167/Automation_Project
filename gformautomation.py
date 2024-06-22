import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchDriverException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

path = r"C:\Users\hp\OneDrive\Documents\project\cat.xlsx"

data = pd.read_excel(path)

name = data['Name'].to_list()
age  = data['Age'].to_list()
gen = data['Gender'].to_list()
dept = data['Gender'].to_list()

try:
    driver = webdriver.Chrome()
    web1 = driver.get("https://forms.gle/zNFpj7PNJGJh5vxn7")
    driver.maximize_window()
except NoSuchDriverException as e:
    print(e)    

time.sleep(3)

try:
    for i in range(len(data)):
      fullname = driver.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input')
      time.sleep(1)
      fullname.clear()
      fullname.send_keys(str(name[i]))

      age_1 = driver.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input')
      time.sleep(1)
      age_1.clear()
      age_1.send_keys(str(age[i])) 

      m = driver.find_element(By.XPATH,'//*[@id="i13"]/div[3]/div')
      f = driver.find_element(By.XPATH,'//*[@id="i16"]/div[3]/div')
      if(gen[i]=="Male"):
          m.click()
      else:
          f.click()
      time.sleep(2)   
    

      
      
      
except NoSuchElementException as f:
    print(f)    

submit = driver.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/div[2]')
driver.execute_script("arguments[0].click()", submit) 
    


time.sleep(2)
driver.close()



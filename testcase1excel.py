import time

from selenium.webdriver.common.by import By

import XLUtilitis
from selenium import webdriver

driver=webdriver.Chrome()
driver.implicitly_wait(10)
driver.get("https://www.instagram.com/")
driver.maximize_window()

path="D:/OneDrive/Excel/login.xlsx"
rows=XLUtilitis.getRowCount(path,'Sheet1')
for r in range(2,rows+1):
    username=XLUtilitis.readData(path,'Sheet1',r,1)
    password=XLUtilitis.readData(path,'Sheet1',r,2)

    driver.find_element(By.XPATH,"//input[@name='username']").send_keys(username)
    driver.find_element(By.XPATH,"//input[@name='password']").send_keys(password)
    driver.find_element(By.XPATH,"//button[@type='submit']").click()

    if driver.title=="Instagram":
        print("Test Passed")
        XLUtilitis.writeData(path,'Sheet1',r,3,"test passed")
    else:
        print('test failed')
        XLUtilitis.writeData(path, 'Sheet1', r, 3, "test failed")

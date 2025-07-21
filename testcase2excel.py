from selenium import webdriver
from selenium.webdriver.common.by import By

import XLUtilitis
driver=webdriver.Chrome()
driver.implicitly_wait(10)
driver.get("https://homeloans.sbi/calculators")
driver.maximize_window()

file='D:/OneDrive/Excel/Loan.xlsx'
rows=XLUtilitis.getRowCount(file,'Sheet1')
for r in range(2,rows+1):

    #fetching data from xlutilitis
    loan_amount=XLUtilitis.readData(file,'Sheet1',r,1)
    tenure=XLUtilitis.readData(file,'Sheet1',r,2)
    intrest_rate=XLUtilitis.readData(file,'Sheet1',r,3)
    emi_amount=XLUtilitis.readData(file,'Sheet1',r,4)

    #passing testdata to application
    driver.find_element(By.NAME,"loanamount").clear()
    driver.find_element(By.NAME, "loanamount").send_keys(loan_amount)

    driver.find_element(By.NAME,"loanTenureValue").clear()
    driver.find_element(By.NAME, "loanTenureValue").send_keys(tenure)

    driver.find_element(By.NAME,"loanInterestRate").clear()
    driver.find_element(By.NAME, "loanInterestRate").send_keys(intrest_rate)

    actual_emi=driver.find_element(By.ID,"totalEmiDef").text
    new_emi=actual_emi.replace("₹"," ").strip()

    if int(emi_amount)==int(new_emi):
        XLUtilitis.writeData(file,'Sheet1',r,6,"Passed")
        XLUtilitis.fill_green(file,'Sheet1',r,6)
    else:
        XLUtilitis.writeData(file, 'Sheet1', r, 6, "Failed")
        XLUtilitis.fill_red(file, 'Sheet1', r, 6)

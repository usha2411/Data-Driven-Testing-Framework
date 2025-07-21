import openpyxl
import XLUtilitis
file="D:/OneDrive/Excel/Project1.xlsx"

#writing data
XLUtilitis.writeData(file,'Sheet1',1,1,"Name")
XLUtilitis.writeData(file,'Sheet1',1,2,"Age")
XLUtilitis.writeData(file,'Sheet1',1,3,"ID")

XLUtilitis.writeData(file,'Sheet1',2,1,"Anaye")
XLUtilitis.writeData(file,'Sheet1',2,2,"5")
XLUtilitis.writeData(file,'Sheet1',2,3,"24514")

#get col row number
print(XLUtilitis.getRowCount(file,'Sheet1'))
print(XLUtilitis.getColumnCount(file,'Sheet1'))

#fill colour by hex code
XLUtilitis.fill_red(file,'Sheet1',1,1)
XLUtilitis.fill_red(file,'Sheet1',1,2)
XLUtilitis.fill_red(file,'Sheet1',1,3)
XLUtilitis.fill_green(file,'Sheet1',2,1)
XLUtilitis.fill_green(file,'Sheet1',2,2)
XLUtilitis.fill_green(file,'Sheet1',2,3)
#read data
XLUtilitis.readData(file,'Sheet1',1,1)
XLUtilitis.readData(file,'Sheet1',1,2)
XLUtilitis.readData(file,'Sheet1',1,3)
XLUtilitis.readData(file,'Sheet1',2,1)
XLUtilitis.readData(file,'Sheet1',2,2)
XLUtilitis.readData(file,'Sheet1',2,3)


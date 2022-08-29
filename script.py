#this is a script to automatically calculate the total credit of dividends
#from the excel sheets


import xlrd


def excelFile():
    filePath = "sbiSample.xlsx"
    workbook = xlrd.open_workbook(filePath)
    worksheet = workbook.sheet_by_index(0)
    matchingString = str(input("input the string to be matched :")).upper()

    # you can change the Column names as per excel sheet
    descriptionColumn = "Description".upper()
    creditColumnName="credit".upper()
    creditInterestStringMatching="CREDIT INTEREST".upper()
        

    columnNumberOfDescription=-1
    columnNumberOfCredit=-1
    fixingRow = -1


    for i in range(worksheet.nrows):
        fixingRow = i
        for j in range(worksheet.ncols):
            if(str(worksheet.cell_value(i,j)).upper()==descriptionColumn):
                columnNumberOfDescription=j
            if(str(worksheet.cell_value(i,j)).upper()==creditColumnName):
                columnNumberOfCredit=j
        if(columnNumberOfCredit>0 and columnNumberOfDescription>0):
            break

    print("rowNumber :"+str(fixingRow)+", descrition :"+str(columnNumberOfDescription)+", credit: "+ str(columnNumberOfCredit))
    totalDividends = 0
    creditInterest = 0
    for i in range(fixingRow+1,worksheet.nrows):
        if matchingString in str(worksheet.cell_value(i,columnNumberOfDescription)).upper():
            totalDividends=totalDividends+worksheet.cell_value(i,columnNumberOfCredit)
        if creditInterestStringMatching in str(worksheet.cell_value(i,columnNumberOfDescription)).upper():
            creditInterest = creditInterest+worksheet.cell_value(i,columnNumberOfCredit)

    print("Total Dividends: "+str(totalDividends))
    print("Total Credit Interest: "+str(creditInterest))

excelFile()




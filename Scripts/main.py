import requests
import json
from tabulate import tabulate
import csv
import openpyxl
from openpyxl import load_workbook

# from openpyxl import Workbook
# from openpyxl.utils import get_column_letter
# from openpyxl import get_highest_row
# from openpyxl import load_workbook

ratios = {'date':'Date', 'symbol':'Ticker', 'grossProfitMargin':'Gross Profit Margin', 'returnOnEquity':'Return on Equity', 'currentRatio':'Current Ratio', 'quickRatio':'Quick Ratio', 'debtEquityRatio':'Debt-Equity Ratio', 'debtRatio':'Debt Ratio', 'priceEarningsRatio':'P/E', 'priceToBookRatio':'P/B'}
income = {'date':'Date', 'period':'Period', 'symbol':'Ticker', 'revenue':'Revenue', 'grossProfit':'Gross Profit', 'operatingIncome':'Operating Income', 'netIncome':'Net Income'}
balance = {'date':'Date', 'period':'Period', 'symbol':'Ticker', 'cashAndCashEquivalents':'Cash and Equivalents', 'totalAssets':'Assets', 'longTermDebt':'Long Term Debt', 'commonStock':'Common Stock', 'retainedEarnings':'Retained Earnings', 'totalDebt':'Total Debt'}
cashflows = {'date':'Date', 'period':'Period', 'symbol':'Ticker', 'debtRepayment':'Debt Repayment', 'commonStockIssued':'Common Stock Issued', 'dividendsPaid':'Dividends Paid', 'freeCashFlow':'Free Cash Flow'}

#returns a value from 1 json
def returnFromJson(dataName, file):
    return file[dataName] #incoming file is in json format

#returns a list of values inputed from 1 json
def returnListFromJson(listOfNames, file):
    temp = []
    for i in listOfNames:
        temp.append(returnFromJson(i, file))
    return temp

def sortMe(dataList, sortOption):
    newDataList = []
    #1. Determine by what item to sort by
    if(sortOption == 0):
        return dataList
    sortOption += 1 #setting the sort option equal to it's index in the respective list
    
    #2. Itereate through given data, and add greatest one by one to new list
    
    for i in range(0, len(dataList)):
        greatestItem = -9999999999 #number is low negative so that any value will be greater than it
        greatestIndex = 0
        indexCount = 0

        for tickerData in dataList:
            #print("Data List: " + str(dataList))
            #print("New Data ListL " + str(newDataList))
            if(float(tickerData[sortOption]) > greatestItem):
                greatestItem = float(tickerData[sortOption])
                greatestIndex = indexCount

            indexCount += 1

        newDataList.append(dataList[greatestIndex])
        del dataList[greatestIndex]
            

    return newDataList


#   Statement Choice, years viewed limit, list of tickers, bool of whether to see it quarterly or not, sort by num
def show(showOption, limit, tickers, quarterly, sortOption):
    #this list will be filled with the headers of the correct financial statement
    headers = []
    
    #determine which statement to view
    dic = ratios
    if(showOption == '2'):
        dic = balance
    elif(showOption == '3'):
        dic = income
    elif(showOption == '4'):
        dic = cashflows

    #load header list
    for key, value in dic.items():
        headers.append(value)
    
    for i in range(limit):
        data = []
        for ticker in tickers:
            dataList = []
            dataList.clear()
            file = getInfo(showOption, ticker, quarterly)
            #if there arent enough finincial statements, return
            if(limit > len(file)):
                print('Ticker ' + ticker.upper() +  'only has ' + str(len(file)) + ' year/quarters of financial statements')
                return
            searchFile = file[i]
            
            for key, value in dic.items():
                dataList.append(searchFile[key])
            data.append(dataList)
        newData = sortMe(data, int(sortOption))
        print(tabulate(newData, headers))
        print('\n')
        data.clear()
            
def create(createOption, limit, tickers, quarterly, sortOption, fileName):
    #this list will be filled with the headers of the correct financial statement
    headers = []

    #Creating New Workbook
    wb = Workbook()
    dest_filename = fileName
    ws1 = wb.active
    workSheetList = ['WorkSheet1', 'WorkSheet2', 'WorkSheet3', 'WorkSheet4', 'WorkSheet5']
    #determine which statement to view
    dic = ratios
    if(createOption == '2'):
        dic = balance
    elif(createOption == '3'):
        dic = income
    elif(createOption == '4'):
        dic = cashflows

    #load header list
    for key, value in dic.items():
        headers.append(value)
    
    for i in range(limit):   
        #Creating Repective Worksheets     
        if(i != 0):
            ws = wb.create_sheet(title=workSheetList[i])
        else:
            ws = wb.active
            ws.title = workSheetList[i]
        #At this point, the current worksheet can be reffered to as ws
        data = []
        for ticker in tickers:
            dataList = []
            dataList.clear()
            file = getInfo(createOption, ticker, quarterly)
            #if there arent enough finincial statements, return
            if(limit > len(file)):
                print('Ticker ' + ticker.upper() +  'only has ' + str(len(file)) + ' year/quarters of financial statements')
                return
            searchFile = file[i]
            
            for key, value in dic.items():
                dataList.append(searchFile[key])
            data.append(dataList)
        newData = sortMe(data, int(sortOption))
        writeToSheet(headers, newData, wb, ws)
        print('\n')
        data.clear()
    #Save workbook in file
    wb.save(filename = dest_filename)

#Update Function
# 1. Update Current Equities to MV
# 2. Add new equity to bottom of list

# 1.
# Go through specified sheet name using /get_sheet_by_name("SheetName") or use .get_active_sheet()
# Ask which sheet, sheets, or all sheets
# Which financial sheet
# Ask Quarterly or Yearly
# Ask to sort
# Collect ticker names from this sheet
# Collect dates from sheet
# Clear Sheet
# Call API for every ticker and search for the specific date and insert values as usual

# docDecision: Which finacial statemet to use, wb: workBook, sortOption, which data point to sort by
def updateEquities(docDecision, wb, sheetNames, sortOption):
    findTickers(wb, sheetNames)

    # #this list will be filled with the headers of the correct financial statement
    # headers = []
    
    # #determine which statement to view
    # dic = ratios
    # if(docDecision == '2'):
    #     dic = balance
    # elif(docDecision == '3'):
    #     dic = income
    # elif(docDecision == '4'):
    #     dic = cashflows

    # #load header list
    # for key, value in dic.items():
    #     headers.append(value)
    
    # for i in range(limit):
    #     data = []
    #     for ticker in tickers:
    #         dataList = []
    #         dataList.clear()
    #         file = getInfo(docDecision, ticker, quarterly)
    #         #if there arent enough finincial statements, return
    #         if(limit > len(file)):
    #             print('Ticker ' + ticker.upper() +  'only has ' + str(len(file)) + ' year/quarters of financial statements')
    #             return
    #         searchFile = file[i]
            
    #         for key, value in dic.items():
    #             dataList.append(searchFile[key])
    #         data.append(dataList)
    #     newData = sortMe(data, int(sortOption))
    #     print(tabulate(newData, headers))
    #     print('\n')
    #     data.clear()

def findTickers(wb, sheetNames):
    tickers = []
    currentSheet = wb[sheetNames[0]]
    maxRow = 1
    #Finding the highest row
    while(currentSheet['B' + str(maxRow)].value is not None):
        maxRow += 1
        if(maxRow > 50):
            break
    #appending tickers from xlsx file to tickers list
    for i in range(2, maxRow):
        tickers.append(currentSheet['B' + str(i)].value)

    #checkTicker returns a list of tickers that didn't work
    badTickers = checkTicker(tickers, True)
    #If there are bad tickers, remove them from the tickers list and inform the user
    if(len(badTickers) > 0):
        for badTicker in badTickers:
            print(badTicker + ' will not be updated')
            tickers.remove(badTicker)
    print(tickers)
    return tickers, badTickers
        


#2.
# Using the following to find the current sheets:
# print(wb2.sheetnames)
#['Sheet2', 'New Title', 'Sheet1']
# Based on years, append to that many sheets
# Use append function to bottom of sheet for each equity


def writeToSheet(headers, data, workBook, workSheet):
    workSheet.append(headers)
    #Data is a list of lists, so info is a single list
    for info in data:
        workSheet.append(info)
    print("Sheet Successfully Created")


#function needs it extract data from a list of json files for the given years and print them out in tabular form
def getInfo(showOption, ticker, quarterly):
    timingPhrase = '' #words that go into url if it is yearly
    showOption = int(showOption)
    ticker = ticker.upper()


    if(quarterly):
        timingPhrase = 'period=quarter&'
    showOption -= 1

    statementList = ['ratios', 'balance-sheet-statement', 'income-statement', 'cash-flow-statement']
    statement = statementList[showOption]

    file = requests.get('https://financialmodelingprep.com/api/v3/' 
                        + statement + '/' 
                        + ticker + '?' 
                        + timingPhrase + '&apikey=0b6e543ddbcbce86657264ae53f7a796')

    return file.json()


#makes sure ticker given is valid
#singleTicker: if true, returns a list of the invalid tickers
def checkTicker(tickers, singleTicker):
    badTickers = [] #only used if singleTicker is ture
    for ticker in tickers:
        passed = False
        ticker = ticker.upper()

        file = requests.get("https://financialmodelingprep.com/api/v3/financial-statement-symbol-lists?apikey=0b6e543ddbcbce86657264ae53f7a796")
        file = file.json()

        for i in file:
            if(ticker == i):
                passed = True
                break
        
        if(passed == False):
            badTickers.append(ticker)
            print('Ticker ' + ticker +  ' is invalid, doesn\'t have financial statements, or is an OTC or ETF')
    if(singleTicker):
        return badTickers
    elif(len(badTicker) > 0):
        return False
    return True

#makes sure input is an int and is within desired range
def checkNum(input, maxNum):
    try:
        input = int(input)
    except ValueError:
        print("Please enter a valid number")
        return False

    if(input > maxNum or input < 1):
        print("Please enter a valid menu option")
        return False
    return True
# -------------------------- Printing -------------------------- #

def printMenu():
    print("------- Actions -------")
    print("1. Show")
    print("2. Create")
    print("3. Update")
    print("4. Help")
    print("5. Exit")
    print('\n')

    temp = input("Please input the action you would like to do: ")
    return temp

def printShowOptions():
    
        print('\n')

        print("------- Options -------")
        print("1. Financial Ratios")
        print("2. Balance Sheet")
        print("3. Income Statement")
        print("4. Statement of Cashflow")
        print('\n')

        temp = input("Please input what you would financial data you would like to use: ")
        return temp

def printShowYears():
    temp = input('How many years/quarters would you like to see? (Max of 5): ')
    return temp

def printShowTicker():
    temp = input('Which ticker/tickers would you like to see? ')
    temp = temp.replace(' ', '')
    temp = temp.split(',')
    return temp

def printYearlyOrQuarterly():
    temp = input('Would you like quarterly numbers or yearly? (Q for Quarter; Y for Yearly): ')
    temp = temp.upper()
    
    if(temp == 'Q'):
        return True
    elif(temp == 'Y'):
        return False
    else:
        print('Invalid Input')
        return -1

def printSort(showOption):
    temp = input('Would you like to sort by any values? (Y/N): ')
    temp = temp.upper()

    if(temp == 'N'):
        return False
    elif(temp == 'Y'):
        if(showOption == '1'):
            printSortOptions(2, ratios)
        elif(showOption == '2'):
            printSortOptions(3, balance)
        elif(showOption == '3'):
            printSortOptions(3, income)
        elif(showOption == '4'):
            printSortOptions(3, cashflows)
        
        temp2 = input("Which value would you like to sort by?")
        return temp2

def printUpdateOptions():
    print("------- Options -------")
    print("1. Update existing Spreadsheet/s too current market value")
    print("2. Add a new ticker to existing spreadsheet/s")
    print('\n')
    temp = input("Please input the action you would like to do: ")
    return temp

def printWhichSheets(wb):
    print("------- Options -------")
    print("1. Enter which sheet or sheets you'd like to update")
    print("2. Update all sheets in workbook")
    print("\n")

    temp = input("Please input the action you would like to do: ")
    
    if(temp == '1'):
        sheetNames = input('Please enter the name of the sheet/s')
        sheetNames = sheetNames.replace(' ', '')
        sheetNames = sheetNames.split(',')
        return sheetNames
    elif(temp == '2'):
        return wb.sheetnames
    else:
        print('Invalid Input')
        return -1

#helper function for printSort
def printSortOptions(skipVal, dic): #skipValue is the index at which afterwards the function should start giving sort options
    i = 1
    j = 1
    for key, value in dic.items():
        if(i <= skipVal):
            i += 1
            continue
        print(str(j) + '. ' + str(value))
        j += 1

def printFileName():

        print('\n')

        temp = input("Please enter the name of your file you'd like to create: ")
        return temp.replace(" ", "")

def jprint(obj):
    # create a formatted string of the Python JSON object
    text = json.dumps(obj.json(), sort_keys=True, indent=4)
    print(text)
# =========================================================================== MAIN LOOP =========================================================================== #
exit = False
i = 0

while not exit:
    print('\n')
    action = printMenu()

    if(not checkNum(action, 5)):
        continue
    # ========== Show ==========
    if(action == '1'):
        #Asks for tickers from user
        tickers = printShowTicker()
        if(not checkTicker(tickers, False)):
            continue
        #Asks for which financial document user wants
        showDecision = printShowOptions()
        if(not checkNum(showDecision, 4)):
            continue
        #How many years or Quarters
        yearDec = printShowYears()
        if(not checkNum(yearDec, 5)):
            continue
        yearDec = int(yearDec)
        #Asks if user wants quarterly or yearly numbers
        quarterly = printYearlyOrQuarterly()
        if(not isinstance(quarterly, bool)):
            continue
        #Ask User if they want to sort
        sort = printSort(showDecision)
        if(isinstance(sort, bool)):
            print('\n')
            show(showDecision, yearDec, tickers, quarterly, 0)
        else:
            print('\n')
            show(showDecision, yearDec, tickers, quarterly, sort)
        

        #function that takes in sho1
        # option and amount of years
        #prints out desired items
    # ========== Create ==========
    elif(action == '2'):
        
        #Asks for tickers from user
        tickers = printShowTicker()
        if(not checkTicker(tickers, False)):
            continue
        #Asks for which financial document user wants
        createDecision = printShowOptions()
        if(not checkNum(createDecision, 4)):
            continue
        #Asks if user wants quarterly or yearly numbers
        quarterly = printYearlyOrQuarterly()
        if(not isinstance(quarterly, bool)):
            continue
        #How many years or Quarters
        yearDec = int(printShowYears())
        if(not checkNum(yearDec, 5)):
            continue
        #What is the file name called
        fileName = printFileName()
        if(fileName.find('.xlsx') != -1):
            continue
        else:
            fileName = fileName + ".xlsx"
        #Ask User if they want to sort
        sort = printSort(createDecision)
        if(isinstance(sort, bool)):
            print('\n')
            create(createDecision, yearDec, tickers, quarterly, 0, fileName)
        else:
            print('\n')
            create(createDecision, yearDec, tickers, quarterly, sort, fileName)
    # ========== Update ==========
    elif(action == '3'):
        #Asks user for what update function they want to use
        updateDecision = printUpdateOptions()
        if(not checkNum(updateDecision, 2)):
            continue

        #Asks for which financial document user wants
        docDecision = printShowOptions()
        if(not checkNum(docDecision, 4)):
            continue

        #Ask for the name of the workbook
        workbookName = input('What is the name of the workbook? ')
        if(workbookName.find('.xlsx') != -1):
            print('')
        else:
            workbookName = workbookName + ".xlsx"
        #Check to see if filename is valid
        try:
            wb = load_workbook(filename = workbookName)
        except (FileNotFoundError, PermissionError) as e:
            print('File Not Found or selected sheet is still open')
            continue
        
        listOfSheets = printWhichSheets(wb)
        if(listOfSheets == -1):
            continue

        #Ask User if they want to sort
        sort = printSort(docDecision)
        if(isinstance(sort, bool)):
            sort = 0
        
        if(updateDecision == '1'):
            updateEquities(docDecision, wb, listOfSheets, sort)
        else:
            print(updateDecision)

    elif(action == '4'):
        print("------- Help Menu -------")
        print("1. Show")
        print("     The show function will print a table containing all of the data you choose to see in the terminal")
        print("     It can show an unlimited amount of tickers, can show quarterly or yearly results, show up to 5 quarters/years at once, and sort by any value")
        print("2. Create")
        print("     The create function will create a new excel sheet based off of the data inputed")
        print("     Each year/quarter will be shown in a different sheet in the workbook")
        print("3. Update")
        print("     The update function has multiple options. It can go through an existing excel sheet created by this program and ")
        print("     update all of the numbers to the most result quarterly/yearly results.")
        print("     The second function is adding a new equity to the bottom of the list for up to 5 years")
        print("4. Help")
        print("5. Exit")
        print('\n')
    elif(action == '5'):
        print("Exiting Program...")
        exit = True

    if(i == 10):
        exit = True
    
    i += 1
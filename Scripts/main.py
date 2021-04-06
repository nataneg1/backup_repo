import requests
import json
from tabulate import tabulate
import csv
import openpyxl

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

#Variables:
#Financial Sheet <-- Covered By Show
#Number of Tickers
#Number of years
#Quarterly or Yearly <-- Covered By Show

#   Statement Choice, years viewed limit, list of tickers, bool of whether to see it quarterly or not, sort by num
def show(showOption, limit, tickers, quarterly, sortOption):
    #this list will be filled with the headers of the correct financial statement
    headers = []
    limit = int(limit)
    
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
            
def create(createOption, tickers, quarterly, sortOption, fileName):
    #this list will be filled with the headers of the correct financial statement
    headers = []
    #limiting spreadsheet creation to 1
    limit = 1
    
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
        writer(headers, newData, fileName)
        print('\n')
        data.clear()


def writer(headers, data, fileName):
    with open (fileName, "w", newline = "") as csvfile:
        sheet = csv.writer(csvfile)
        sheet.writerow(headers)
        for x in data:
            sheet.writerow(x)
    print("File Successfully Created")


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
def checkTicker(tickers):

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
            print('Ticker ' + ticker +  ' is invalid, doesn\'t have financial statements, or is an OTC or ETF')
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
    print("3. IDK Yet")
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
        tickers = printShowTicker()
        if(not checkTicker(tickers)):
            continue

        showDecision = printShowOptions()
        if(not checkNum(showDecision, 4)):
            continue

        yearDec = printShowYears()
        if(not checkNum(yearDec, 5)):
            continue

        quarterly = printYearlyOrQuarterly()
        if(not isinstance(quarterly, bool)):
            continue
        
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
        
        tickers = printShowTicker()
        if(not checkTicker(tickers)):
            continue

        createDecision = printShowOptions()
        if(not checkNum(createDecision, 4)):
            continue

        quarterly = printYearlyOrQuarterly()
        if(not isinstance(quarterly, bool)):
            continue

        fileName = printFileName()
        fileName = fileName + ".csv"

        sort = printSort(createDecision)
        if(isinstance(sort, bool)):
            print('\n')
            create(createDecision, tickers, quarterly, 0, fileName)
        else:
            print('\n')
            create(createDecision, tickers, quarterly, sort, fileName)

    elif(action == '3'):
        print("IN IDK YET")
    elif(action == '4'):
        print("IN HELP")
    elif(action == '5'):
        print("Exiting Program...")
        exit = True

    if(i == 10):
        exit = True
    
    i += 1
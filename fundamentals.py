from datetime import datetime, timedelta
import pandas as pd
import re

# creates fixed pd dataFramed
def createFrame(num, col1, col2, datatype):
    # empty data Frame to store summary of invoices 
    df = pd.DataFrame()
    
    # creating fixed number of columns
    for i in range(num):
        df.insert(i, "{}".format(i), ['0'], True)
    
    # correct datatype
    dataTypeCor(df,  col1, col2, datatype)
    
    return df

# checks excel file
def allowed_file(filename):
    allowed_extensions = set(['xlsx'])
    return "." in filename and \
        filename.rsplit('.',1)[1].lower() in allowed_extensions

# change datatype for columns
def dataTypeCor(df, st, fin, dataType):
    """ 
    df = dataFrame
    cols = range of columns 
    dataType = dtype 
    """ 
    
    df[df.columns[st:fin]] = df[df.columns[st:fin]].astype(dataType)
    return None

# formats to currency
def money(num):
    for i in range(len(num)):
        if num[i] != '-':
             num[i] = "{:,.2f}".format(num[i])
        # if num[i] == 0:
        #     num[i] = "-"
        # else: 
        #     num[i] = "{:,.2f}".format(num[i])
        
# change date format dd/mm/yyyy
def dateFormat(dates):
    for i in range(len(dates)):
        #Payable upto 31-MAY-2022 
        dates[i] = "Payable upto {}".format(dates[i].strftime('%d-%b-%Y'))
            

# calculate interest per month
def getInterest(item, currentDate, schpay, interest, dates):
    """ 
    currentDate previous date 
    schpay - predefined amount
    dates - array to store string dates for all due dates
    difference - days difference from initial and due date
    """ 
    prevDiff = 0
    for i in range(0, 11):

        if i == 0:
            dates[i] = geom(currentDate)
            dffDays =  int(dates[0].strftime('%d')) - int(currentDate.strftime('%d') )
            prevDiff = dffDays
            item[23 + i] = round(schpay + schpay*prevDiff*interest/100,2)
        
        else:
            dates[i] = geom(dates[i-1] + timedelta(days=1))
            d1 = datetime.strptime(str(dates[i]), "%Y-%m-%d")
            d2 = datetime.strptime(str(dates[i-1]), "%Y-%m-%d")
            dffDays = (d1 - d2).days
            prevDiff = dffDays + prevDiff
            item[23 + i] = round(schpay + schpay*prevDiff*interest/100, 2)
            
    

# find date end of the month
def geom(date):
    #extract month, day and year
    month = int(date.strftime('%m'))
    year = int(date.strftime('%Y'))
    
    day = 1  
    if month == 12:
        month = 1
        year = year + 1
    else:
        month = month + 1
    
    
    newDate = "{}/{}/{}".format(day, month, year)
    newDate = datetime.strptime(newDate, '%d/%m/%Y').date()
    newDate = newDate - timedelta(days=1)
    
    return newDate

# validate email ID
def isValid(email):
    regex = regex = re.compile(r'([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')
    if re.fullmatch(regex, email):
      return True
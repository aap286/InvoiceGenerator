from fundamentals import dataTypeCor, money, getInterest, dateFormat
from flask import Flask, request, render_template
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
import os, win32com.client, winshell
import pandas as pd
import numpy as np
import warnings
import webview
import pdfkit
import pythoncom



# suppresses warnings 
warnings.simplefilter(action='ignore', category=FutureWarning)

# location of html to pdf converter
sysmConfig = pdfkit.configuration(wkhtmltopdf='style\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')

# import configuration file 
configFrame = pd.read_csv("0_Configuration\\configuration.csv", header=None, names=['Name', 'Value']).T # transposes dataframe

# renaming columns
configFrame.columns = configFrame.iloc[0] # column names matches to the first row  
configFrame = configFrame.drop(configFrame.index[0]) # removing the first row

# correcting dataType
dataTypeCor(configFrame, 0,11, float)

# config data types
""" 
Maintenance Rate                    float64
Weighted Interest                   float64
Reserve Fund                        float64
Recovery Interest rate              float64
Non occupancy charges               float64
Payment Discount                    float64
Payment Discount Period             float64
Delayed Payment charges             float64
CGST Rate                           float64
SGST Rate                           float64
Delayed Payment charges Interest    float64
Recovery Excess Interest            string
NA Tax arrears                       object
Deemed Conveyance arrears            object
Painting Project arrears             object
Lift Project arrears                 object
Delay Payment Charges                object
emailPasscode                        object
email Address                        object
""" 

# converts vertical data frame into list which each element is a dictionary 
configFrame = list(configFrame.to_dict().values())

# print(configFrame)

# empty array to store configFrame values
config = [None]
for value in configFrame:
    config.append(value['Value'])

# memory management
del configFrame   

# store row names
itemName = [""]*36 

# store rate values
itemRate = [""] * 36

# *** SECTION NAMES LIST
itemSectionNames = [''] * 5
    
# send mail
email_password = config[17]

# sender email address
sender = config[18]

# row 3 is it less or add
recoveryInterestRate = 'Add' if config[12].upper() == 'ADD' else 'Less'
  

# flask web service        
def create_app():

    app = Flask(__name__, template_folder= 'style\\templates', static_folder='style') 
        
    # home page
    @app.route('/', methods=['GET','POST'])
    def home():  
      
        # after submitting the form
        if request.method == 'POST':
            
            # ? FORM VALUES
            file = request.files["file"] # input excel file
            invoiceDate = request.form.get('invoiceDate') # invoice creation date
            year = request.form.get('year') # invoice year range
            period = request.form.get('period') # invoice period name 
            resolutionText = request.form.get('resolutionText') # invoice resolution
            # subject = request.form['subject']
            # body = request.form['message']
       
            # accepts excel file
            filename = secure_filename(file.filename) # removes / from file name
            filename = os.path.join('Input', filename)  # directory of input files
            file.save(filename) # saves file in input folder


            # ? CONVERTS FORMAT DATE
            invoiceDate = datetime.strptime(invoiceDate, '%Y-%m-%d').date()

            #? DUE DATE FROM THE INVOICE GENERATION
            # reinstate due date of invoice
            invoiceDateDue = invoiceDate + timedelta(days=config[7])

            # change format of invoice  date
            invoiceDateStr = invoiceDate.strftime('%d/%m/%Y')
            invoiceDateDueStr = invoiceDateDue.strftime('%d-%b-%Y')
                
            # directory of saved invoice pdfs
            dirName = '{}\\0_Invoices\\{}'.format(os.getcwd(), year)
                
            # creates directory 
            try:
                os.makedirs(dirName)
                    
                # creates shortcut
                desktop = winshell.desktop()
                path = os.path.join(desktop, 'Invoices.lnk'.format(os.getcwd()))
                target = '{}\\0_Invoices'.format(os.getcwd())
                icon = '{}\\style\\icon\\invoice.ico'.format(os.getcwd())
                shell = win32com.client.Dispatch("WScript.Shell", pythoncom.CoInitialize())
                shortcut = shell.CreateShortCut(path)
                shortcut.Targetpath = target
                shortcut.IconLocation = icon
                shortcut.save()
            except FileExistsError:
                None
                
            # import form excel file
            apartments = pd.read_excel(filename, "Sheet1")   

            # ????? SECTION NAMES FOR A1, A2, B, AND C
            itemSectionNames[1] = 'Current Period (FY{})'.format(year) 
            itemSectionNames[2] = 'Arrears from previous periods/invoices (if any)'
            itemSectionNames[3] = 'If Paid Before {}'.format(invoiceDateDueStr)
            itemSectionNames[4] = "Payment Schedule if Paid After " + invoiceDateDueStr 
                
            # ????? ROW NAMES FOR SECTION A1 & A2,B,C
            # TODO SECTION A1
            itemName[1] = "Maintenance @ Rs.{}/sqft/mth".format(config[1])
            itemRate[1] = str(config[1])

            itemName[2] = "Less: Interest credit on Corpus @ {}% p.a.".format(config[2])
            itemRate[2] = "{}%".format(config[2])
            
            itemName[3] = """{}: Recovery of excess interest credit @ {}% on   
                                    corpus (in FY {}) as per {}""".format(recoveryInterestRate, config[4], year, resolutionText)
            itemRate[3] = "{}%".format(config[4])

            itemName[4] = "Net Maintenance Payable"

            if recoveryInterestRate == 'Add':
                itemRate[4] = "(1-2+3)"
            else:
                itemRate[4] = "(1-2-3)"

            itemName[5] = "Reserve Fund @ Rs.{}/sqft/mth (excl GST)".format(config[3])
            itemRate[5] = str(config[3])

            itemName[6] = "Non-Occupancy Charges @ {}% of item 1 (if rented out)".format(config[5])
            itemRate[6] = "(See note-1)"

            itemName[7] = "CGST @ {}% (on items 5,6)".format(config[9])
            itemName[8] = "SGST @ {}% (on items 5,6)".format(config[10])
            # * itemRate column values are empty !!!!!!!!!

            itemName[9] = "Total Current Period Dues (A1)"
            itemRate[9] = "(4+5+6+7+8)"

            # TODO SECTION A2
            itemName[10] = "Mntce Arrears from previous periods invoices (if any)"
            itemName[11] = "NA Tax arrears"
            itemName[12] = "Deemed-Conveyance arrears"
            itemName[13] = "Painting Project arrears"
            itemName[14] = "Lift Project arrears"
            itemName[15] = "Delay Payment Charges @ {}% (on item nos. 10 to 14)".format(config[8])
            # * itemRate column values are empty !!!!!!!!!

            itemName[16] = "CGST @ {}% (on items 15)".format(config[9])
            itemRate[16] = "{}%".format(config[9])

            itemName[17] = "SGST @ {}% (on items 15)".format(config[10])
            itemRate[17] = "{}%".format(config[9])

            itemName[18] = 'Total Arrears (A2)'
            itemRate[18] = 'Sum (10 to 17)'

            itemName[19] = "Grand Total Due (Payable upto {}) (A1 + A2)".format(invoiceDateDue.strftime('%d-%b-%Y'))
            itemRate[19] = '(9+18)'

            # TODO SECTION B
            itemName[20] = "Full Year Amount"
            itemName[21] = "Less: Discount @ {}% (on item 4)".format(config[6])
            itemName[22] = "Net Payble on or before {}".format(invoiceDateDue.strftime('%d-%b-%Y'))
            itemRate[21] = itemRate[22] = "{}%".format(config[6])

           
            # ? Iterating through each user
            for j in range(len(apartments)):
                    
                    # pivoits around each user
                    userOne = apartments.loc[j] 

                    # output of userOne
                    """ 
                    S.NO                          1
                    B.NO                         A1
                    Flat No.                    101
                    Area, Sq.ft               2207
                    Actual Corpus Deposit    242770
                    """
                    
                    # store amount Rs
                    item = ['-']*35
                    
                    # store delayed payment for arears
                    # delayedItem = [0]*11
                    
                    # name of user
                    name = userOne['Name'] 
                    # apartment and building number
                    aptNo = "{a}/{b}".format(a=userOne['B.NO'], b=userOne['Flat No.']) # apartment number concanted
                    # status of occupancy
                    status = (
                        'Self-Occupied' 
                        if userOne.get('Self-occupied', '').lower() == 'y' 
                        else 'Rented'
                    )
     
                    # corpus value
                    corpus = "{:,}".format(userOne['Actual Corpus Deposit'])
                    area = "{:,}".format(userOne["Area, Sq.ft"]) 


                    # Section A
                    # maintenance @ 1.80/sqft/mth
                    item[1] = round(config[1] * 12 * userOne["Area, Sq.ft"], 0)

                    # Less interest credit on Corpus
                    item[2] = round(config[2] * userOne["Actual Corpus Deposit"] / 100 ,0)

                    # Recovery excess interest rate
                    item[3] = round(config[4] * userOne["Actual Corpus Deposit"] / 100, 0)
                    
                    # Net Maintenance Payable
                    if item[3] == 'Add':
                        item[4] = round(item[1] - item[2] + item[3], 0)
                    else:
                        item[4] = round(item[1] - item[2] - item[3], 0)


                    # Reserve Fund
                    item[5] = round(config[3] * 12 * userOne["Area, Sq.ft"], 0)

                    # Non-Occupancy Charges
                    if status.lower() == 'rented':
                        item[6] = round(config[5] * item[1] / 100, 0) 
                    else:
                        item[6] = 0
                
                    # CGST & SGST 
                    item[7] = round( (item[5] + item[6]) * config[9] / 100, 0)
                    item[8] = round( (item[5] + item[6]) * config[10] / 100, 0)

                    # Total Current period dues (A1) (4+5+6+7+8)
                    item[9] = round(item[4] + item[5] + item[6] + item[7] + item[8], 0)

                    # for rows 10,12,12,13,14
                    for i in range(10, 15):
                        
                        if config[i+3].lower() == 'y' and isinstance(userOne[i-3], np.int64):
                            item[i] = userOne[i-3] if userOne[i-3]  and userOne[i-3] > 0 else 0.00
                            
                        else:
                            item[i] = 0.0
                    
                    if config[17].lower() == 'y':
                        item[15] = round( (item[10] + item[11] + item[12] + item[13] + item[14]) * config[8] / 100 ,0)
                    
                    if item[15] != '-':
                        item[16] = round(item[15] * config[9] / 100 ,0)
                        item[17] = round(item[15] * config[10] / 100 ,0)
                        item[18] = sum(item[11:18])

                    # GRAND TOTAL A1 + A2 && FULL YEAR AMT
                    item[19] = item[20] = item[9] + item[18]

                    # less discount on item 4
                    item[21] = round( item[4] * config[6] / 100, 0)

                    # net payable on or before
                    item[22] = round(item[20] - item[21], 0)

                    # schedule payment after due date
                    schpay = item[20] #round(item[4] + item[5] + item[6] + (item[5] + item[6]) * config[7] / 100, 0)

                    # # interest per day
                    interest = config[11] / 365

                    # dates array 
                    dates = [0] * 12
                    
                    # generates payment plan for due dates entire year
                    getInterest(item, invoiceDateDue, schpay, interest, dates)
                
                    # reformat numbers to currency
                    money(item)
                    
                    # reformat date 
                    dateFormat(dates)
                    itemName[23:36] = dates
                    

                    # indexing invoice docs
                    index = "{}/{}".format(str(101+j), year)


                    item[2] = "({})".format(item[2])
                    item[3] = "({})".format(item[3]) if recoveryInterestRate == 'Less' else item[3]
                    item[6] = "-" if status.lower() != 'rented' else item[6]

                    # !!!! REMOVE CONFIG, EMAIL SUBJECT & BODY
                    html = render_template('output.html',
                            year = year,period = period, 
                            index=index, invDate = invoiceDateStr, 
                            name = name, aptNo = aptNo, status = status, area = area,
                            corpus = corpus, 
                            itemName = itemName, itemRate = itemRate, item = item, itemSectionNames = itemSectionNames
                            ) 
                    
                    # PDF options
                    options = {
                            "orientation": "portrait",
                            "page-size": "A4",
                            "margin-top": "1.0cm",
                            "margin-right": "1cm",
                            "margin-bottom": "1.0cm",
                            "margin-left": "1cm",
                            "encoding": "UTF-8",
                            "enable-local-file-access": ""
                    }
                        
                    
                    # pdf file name 
                    aptNoSave = "{a} {b}".format(a=userOne['B.NO'], b=userOne['Flat No.'])
                    pdfkit.from_string(html, '0_Invoices\\{}\\{}.pdf'.format(year,aptNoSave), options=options, 
                                       configuration=sysmConfig, css=['style\\css\\outputstyle.css'])
                    
            return render_template('progress.html')
                
    
        # template before running the file        
        return render_template('home.html')

    if __name__ == "__main__":
        # app.run(debug=True)
        webview.create_window('Invoice',app,
                width=850, height=750)
        webview.start()

# runs service
create_app()     
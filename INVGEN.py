from fundamentals import dataTypeCor, createFrame, allowed_file, money, getInterest, dateFormat, isValid
from flask import Flask, request, render_template, make_response
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
from email.message import EmailMessage
import os, win32com.client, winshell
import pandas as pd
import numpy as np
import warnings
import smtplib
# import webview
import pdfkit
import math
import ssl


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

# empty array to store configFrame values
config = []
for value in configFrame:
    config.append(value['Value'])

# memory management
del configFrame     

# store row names
itemName = [""]*31 
    
# send mail
email_password = config[16]

# sender email address
sender = config[17]
        

# flask web service        
def create_app():

    app = Flask(__name__, template_folder= 'style\\templates', static_folder='style') 
    # window = webview.create_window('Invoice', app,
    #             width=850, height=750)
    
    # home page
    @app.route('/', methods=['GET','POST'])
    def home():  
        
        # create frame of fixed size
        # df = createFrame(41, 12 ,41, float)   
        
        # after submitting the form
        if request.method == 'POST':
            
            file = request.files["file"] # input excel file
            invoiceDate = request.form.get('invoiceDate') # invoice creation date
            year = request.form.get('year') # invoice year range
            period = request.form.get('period') # invoice period name
            
            subject = request.form['subject']
            body = request.form['message']
            
            # TRUE when no inputs are empty
            # otherwise refreshes page
            if invoiceDate != '' and year != '' and period != '' and file != '' and  allowed_file(file.filename):
            # subject != '' and body != '' and

                # accepts excel file
                filename = secure_filename(file.filename) # removes / from file name
                filename = os.path.join('Input', filename)  # directory of input files
                file.save(filename) # saves file in input folder
                
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
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shortcut = shell.CreateShortCut(path)
                    shortcut.Targetpath = target
                    shortcut.IconLocation = icon
                    shortcut.save()
                except FileExistsError:
                    None
                
                # import form excel file
                apartments = pd.read_excel(filename, "Sheet1")    
                
                # convert format of date
                invoiceDate = datetime.strptime(invoiceDate, '%Y-%m-%d').date()

                # reinstate due date of invoice
                invoiceDateDue = invoiceDate + timedelta(days=config[6])
                
                # adding columns to summary excel file
                itemName[1] = "Maintenance @ Rs.{}/sqft/mth".format(config[0])
                itemName[2] = "Less: Interest credit on Corpus @ {}% p.a.".format(config[1])
                itemName[3] = "Add: Recovery of excess interest credit @ {}% on corpus (in FY {}) as per 17th AGM Resolution dtd 19.12.2021".format(config[3], year)
                itemName[4] = "Net Maintenance Payable"
                itemName[5] = "Reserve Fund @ Rs.{}/sqft/mth (excl GST)".format(config[2])
                itemName[6] = "Non-Occupancy Charges @ {}% of item 1 (if rented out)".format(config[4])
                itemName[12] = "Delay Payment Charges @ {}% (on item nos. 8 to 11)".format(config[7])
                itemName[13] = "Total Payment"
                itemName[14] = "CGST @ {}% (on items 5, 6, 8, 9 & 12)".format(config[8])
                itemName[15] = "CGST @ {}% (on items 5, 6, 8, 9 & 12)".format(config[9])
                itemName[16] = "Grand Total Due (Payable upto {})".format(invoiceDateDue.strftime('%d-%b-%Y'))
                itemName[17] = "Full Year Amount"
                itemName[18] = "Less: Discount @ {}%".format(config[5])
                itemName[19] = "Net payable on or before {}".format(invoiceDateDue.strftime('%d-%b-%Y'))
                
                # read line by line for each user
                for j in range(len(apartments)):
                    
                    # extract first row from apartments
                    userOne = apartments.loc[j] 

                    # output
                    """ 
                    S.NO                          1
                    B.NO                         A1
                    Flat No.                    101
                    Area, Sq.fit               2207
                    Actual Corpus Deposit    242770
                    """
                    
                    # store amount
                    item = [0]*31
                    
                    # store delayed payment for arears
                    delayedItem = [0]*11
                    
                    # store due payment grand total
                    delayedTotal = [0]*11

                    # name of user
                    name = userOne['Name'] 

                    # apartment and building number
                    aptNo = "{a}/{b}".format(a=userOne['B.NO'], b=userOne['Flat No.']) # apartment number concanted

                    # status of occupancy
                    status = "Rented" 
                    try:
                        userOne['Self-occupied'].lower() == "y"
                        status = 'Self-Occupied'
                    except:
                        None
                        
                    # corpus value
                    corpus = "{:,}".format(userOne['Actual Corpus Deposit'])  


                    # Section A
                    # maintenance @ 1.80/sqft/mth
                    item[1] = round(config[0] * 12 * userOne["Area, Sq.fit"], 2)

                    # Interest credit
                    item[2] = round(config[1] * userOne["Actual Corpus Deposit"] / 100 ,2)

                    # Recovery excess interest rate
                    item[3] = round(config[3] * userOne["Actual Corpus Deposit"] / 100, 2)
                    
                    # Net Maintenance Payable
                    item[4] = round(item[1] - item[2] + item[3], 2)

                    # Reserve Fund
                    item[5] = round(config[2] * 12 * userOne["Area, Sq.fit"], 2)

                    # Non-Occupancy Charges
                    try:
                        status == 'Rented'
                        item[6] = round(config[4] * item[1] / 100, 2) 
                    except:
                        None

                    # checks if user has value 
                    for i in range(7,12):
                        try:
                            if not math.isnan(userOne[i]): 
                                item[i] = userOne[i]
                                itemName[i] = str(userOne.index[i])
                        except:
                            None
                    
                    # delayed payment charges
                    item[12] = round( (item[8] + item[9] + item[10] + item[11]) * config[10] / 100 ,2)
                    
                    # Total Payment
                    item[13] = round(item[4] + item[5] + item[6] + item[7] + item[8] + item[9] + item[10] + item[11] + item[12], 2)
                    
                    # CGST & SGST
                    item[14] = item[5] + item[6]
                    item[15] = item[5] + item[6]
            
            
                    # additional charges
                    for i in range(11,16):
                        try:
                            if config[i].lower() == 'y':
                                item[14] = item[14] + item[i-3]
                        except:
                            None
                    
                    item[14] = round( item[14] *config[8] / 100 , 2)
                    item[15] = round( item[15] * config[9] / 100 , 2)

                    # Grand Total
                    item[16] = round(item[13] + item[14] + item[15] , 2)

                    # Section B
                    # full year amount
                    item[17] = item[16]

                    # discount @ 2.5%
                    item[18] = round(config[5] * item[4] /100 , 2)

                    # Net payable before May 15 th
                    item[19] = round(item[17] - item[18], 2)

                    # schedule payment after due date
                    schpay = round(item[4] + item[5] + item[6] + (item[5] + item[6]) * config[7] / 100, 2)

                    # interest per day
                    interest = config[7] / 365

                    # dates array 
                    dates = [0] * 11
                    
                    # generates payment plan for due dates entire year
                    getInterest(item, invoiceDateDue, schpay, interest, dates)
                
                    # reformat numbers to currency
                    money(item)
                    money(delayedItem)
                    
                    # reformat date 
                    dateFormat(dates)
                    itemName[20:31] = dates
                    
                    emailBody = None
                    emailSubject = None
                    
                    # change format of invoice  date
                    invoiceDateStr = invoiceDate.strftime('%d/%m/%Y')
                    invoiceDateDueStr = invoiceDateDue.strftime('%d-%b-%Y')
                    
                    # indexing invoice docs
                    index = "{}/{}".format(str(101+j), year)
                    
                    html = render_template('output.html', item = item, config = config, itemName=itemName,
                        delayedItem=delayedItem, delayedTotal=delayedTotal,                  
                        name=name, aptNo = aptNo, status=status,corpus=corpus, index=index, 
                        invoiceDate=invoiceDateStr, invoiceDateDue=invoiceDateDueStr,
                        emailBody = emailBody, emailSubject = emailSubject,
                        year=year, period=period) 
                    
                    # PDF options
                    options = {
                            "orientation": "portrait",
                            "page-size": "A4",
                            "margin-top": "2.0cm",
                            "margin-right": "0.5cm",
                            "margin-bottom": "1.0cm",
                            "margin-left": "0.5cm",
                            "encoding": "UTF-8",
                            "enable-local-file-access": ""
                    }
                        
                    
                    # pdf file name 
                    aptNoSave = "{a} {b}".format(a=userOne['B.NO'], b=userOne['Flat No.'])
                    pdfkit.from_string(html, '0_Invoices\\{}\\{}.pdf'.format(year,aptNoSave), options=options, 
                                       configuration=sysmConfig, css=['style\\css\\outputstyle.css'])
                    
                    # convert int to float
                    for t in range(len(item)):
                        try:
                            item[t] = float(item[t])
                        except:
                            None
                    
                    # inserting row to summary data Frame
                    # row = np.concatenate((userOne[1:11], item))
                    # df = df.append(pd.Series(row, index=df.columns[:len(row)]), ignore_index=True)
                                        
                    # email Object
                    msg = EmailMessage()
                    msg['From'] = config[17]
                    msg['To'] = userOne[12]
                    msg['Subject'] = subject
                    
                    # forming body of the email
                    salution = "Dear {},\n".format(name)
                    msg.set_content(salution + body)
                    
                    
                    # user has email
                    # try:
                    #     if userOne[12] != '' and isValid(userOne[12]):
                        
                    #         # opens pdf file and attaches to draft
                    #         with open('0_Invoices\\{}\\{}.pdf'.format(year,aptNoSave), 'rb') as content_file:
                    #             content = content_file.read()
                    #             msg.add_attachment(content, maintype='application', subtype='pdf', filename='example.pdf')
                        
                    #         # creates a safe network
                    #         context = ssl.create_default_context()
                    #         with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
                    #             smtp.login(sender, email_password)
                    #             smtp.sendmail(sender, userOne[12], msg.as_string())
                    # except:
                    #     None
                        
                # drop first row
                # df.drop(index=df.index[0], axis=0, inplace=True)   
                
                # renaming columns
                # for i in range(11):
                #     df.rename(columns = {df.columns[i]:apartments.columns[i+1]}, inplace = True)     
                # for k in range(30):
                #     df.rename(columns = {df.columns[k+11]:itemName[k+1]}, inplace = True) 
                       
                # exporting summary data Frame
                # df.to_csv('0_Invoices\\{}\\1.csv'.format(year))
                
                return render_template('progress.html')
                
    
        # template before running the file        
        return render_template('home.html')

    if __name__ == "__main__":
        app.run(debug=True)
        # webview.start()

# runs service
create_app()

        
from fundamentals import dataTypeCor, money, getInterest, dateFormat
from flask import Flask, request, render_template
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
import os, win32com.client, winshell
import pandas as pd
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
config = [None]
for value in configFrame:
    config.append(value['Value'])

# memory management
del configFrame   

# store row names
itemName = [""]*35 

# store rate values
itemRate = [""] * 35

# *** SECTION NAMES LIST
itemSectionNames = [''] * 5
    
# send mail
email_password = config[17]

# sender email address
sender = config[18]
        

# flask web service        
def create_app():

    app = Flask(__name__, template_folder= 'style\\templates', static_folder='style') 
    window = webview.create_window('Invoice', app,
                width=850, height=750)
    
    # home page
    @app.route('/', methods=['GET','POST'])
    def home():  
        # TODO: HAVE TO WORK ON ITEM (AMOUNT AND COLUMN NAMES FOR SECTION C)
        # create frame of fixed size
        # df = createFrame(41, 12 ,41, float)   
        
        # after submitting the form
        if request.method == 'POST':
            
            # ? FORM VALUES
            file = request.files["file"] # input excel file
            invoiceDate = request.form.get('invoiceDate') # invoice creation date
            year = request.form.get('year') # invoice year range
            period = request.form.get('period') # invoice period name 
            subject = request.form['subject']
            body = request.form['message']
       
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

            itemName[2] = "Less: Interest credit on Corpus @ {}% p.a.".format(config[1])
            itemRate[2] = "{}%".format(config[2])
            
            itemName[3] = """Add: Recovery of excess interest credit @ {}% on   
                                    corpus (in FY {}) as per 17th AGM Resolution dtd 19.12.2021""".format(config[4], year)
            itemRate[3] = "{}%".format(config[4])

            itemName[4] = "Net Maintenance Payable"
            itemRate[4] = "(1-2+3)"

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
            itemName[15] = "Delay Payment Charges @ {}% (on item nos. 8 to 11)".format(config[8])
            # * itemRate column values are empty !!!!!!!!!

            itemName[16] = "CGST @ {}% (on items 15)".format(config[9])
            itemRate[16] = "{}%".format(config[9])

            itemName[17] = "SGST @ {}% (on items 15)".format(config[10])
            itemRate[17] = "{}%".format(config[9])

            itemName[18] = 'Total Arrears (A2)'
            itemRate[18] = 'Sum(10 to 17)'

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
                    
                    # store due payment grand total
                    delayedTotal = [0]*11

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


                    # Section A
                    # maintenance @ 1.80/sqft/mth
                    item[1] = round(config[1] * 12 * userOne["Area, Sq.ft"], 2)

                    # Less interest credit on Corpus
                    item[2] = round(config[2] * userOne["Actual Corpus Deposit"] / 100 ,2)

                    # Recovery excess interest rate
                    item[3] = round(config[4] * userOne["Actual Corpus Deposit"] / 100, 2)
                    
                    # Net Maintenance Payable
                    item[4] = round(item[1] - item[2] + item[3], 2)

                    # Reserve Fund
                    item[5] = round(config[2] * 12 * userOne["Area, Sq.ft"], 2)

                    # Non-Occupancy Charges
                    try:
                        status == 'Rented'
                        item[6] = round(config[5] * item[1] / 100, 2) 
                    except:
                        None

                    # CGST & SGST 
                    item[7] = round( (item[5] + item[6]) * config[9] / 100, 2)
                    item[8] = round( (item[5] + item[6]) * config[10] / 100, 2)

                    # Total Current period dues (A1) (4+5+6+7+8)
                    item[9] = round(item[4] + item[5] + item[6] + item[7] + item[8], 2)

                    # GRAND TOTAL A1 + A2 && FULL YEAR AMT
                    item[19] = item[20] = item[9] #+ item[18]

                    # less discount on item 4
                    item[21] = round( item[4] * config[6] / 100, 2)

                    # net payable on or before
                    item[22] = round(item[20] - item[21], 2)

                    # # checks if user has value 
                    # # for i in range(7,12):
                    #     try:
                    #         if not math.isnan(userOne[i]): 
                    #             item[i] = userOne[i]
                    #             itemName[i] = str(userOne.index[i])
                    #     except:
                    #         None
                    
                    # # delayed payment charges
                    # item[12] = round( (item[8] + item[9] + item[10] + item[11]) * config[10] / 100 ,2)
                    
                    # # Total Payment
                    # item[13] = round(item[4] + item[5] + item[6] + item[7] + item[8] + item[9] + item[10] + item[11] + item[12], 2)
                    
                    # # CGST & SGST
                    # item[14] = item[5] + item[6]
                    # item[15] = item[5] + item[6]
            
            
                    # # additional charges
                    # for i in range(11,16):
                    #     try:
                    #         if config[i].lower() == 'y':
                    #             item[14] = item[14] + item[i-3]
                    #     except:
                    #         None
                    
                    # item[14] = round( item[14] *config[8] / 100 , 2)
                    # item[15] = round( item[15] * config[9] / 100 , 2)

                    # # Grand Total
                    # item[16] = round(item[13] + item[14] + item[15] , 2)

                    # # Section B
                    # # full year amount
                    # item[17] = item[16]

                    # # discount @ 2.5%
                    # item[18] = round(config[5] * item[4] /100 , 2)

                    # # Net payable before May 15 th
                    # item[19] = round(item[17] - item[18], 2)

                    # schedule payment after due date
                    schpay = item[20] #round(item[4] + item[5] + item[6] + (item[5] + item[6]) * config[7] / 100, 2)

                    # # interest per day
                    interest = config[11] / 365

                    # dates array 
                    dates = [0] * 11
                    
                    # generates payment plan for due dates entire year
                    getInterest(item, invoiceDateDue, schpay, interest, dates)
                
                    # reformat numbers to currency
                    money(item)
                    # money(delayedItem)
                    
                    # reformat date 
                    dateFormat(dates)
                    itemName[23:35] = dates
                    

                    # indexing invoice docs
                    index = "{}/{}".format(str(101+j), year)
    
                    # !!!! REMOVE CONFIG, EMAIL SUBJECT & BODY
                    html = render_template('output.html',
                            year = year,period = period, 
                            index=index, invDate = invoiceDateStr, 
                            name = name, aptNo = aptNo, status = status,
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
                    
                    # convert int to float
                    for t in range(len(item)):
                        try:
                            item[t] = float(item[t])
                        except:
                            None
                    
                    # inserting row to summary data Frame
                    # row = np.concatenate((userOne[1:11], item))
                    # df = df.append(pd.Series(row, index=df.columns[:len(row)]), ignore_index=True)
                                        
                    # # email Object
                    # msg = EmailMessage()
                    # msg['From'] = config[17]
                    # msg['To'] = userOne[12]
                    # msg['Subject'] = subject
                    
                    # # forming body of the email
                    # salution = "Dear {},\n".format(name)
                    # msg.set_content(salution + body)
                    
                    
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
        # app.run(debug=True)
        webview.start()

# runs service
create_app()

        
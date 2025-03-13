import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pyodbc
import re
from cryptography.fernet import Fernet
import os
from os import walk
import logging
import sys
import traceback
import pandas as pd
import glob

class Utility:

    from time import strftime
    from datetime import date, timedelta
#
#    conn = pyodbc.connect('Driver={SQL Server};Server=c4w32097.itcs.hpecorp.net\SQLExpress;Database=FINANCE_CTS_RPA;UID=sa;Pwd=GFSq@123')
#    key = b'pRmgMa8T0INjEAfksaq2aafzoZXEuwKI7wDe4c1F8AY='

    def get_config(self, configfile= "config.ini"):
        import configparser
        config= configparser.ConfigParser()
        config.read(configfile)
        return config

    def GetFiscalYr(self,date):
        if date.today().month <= 10:
            GetFiscalYr = date.today().year
        else:
            GetFiscalYr = date.today().year+1
        return GetFiscalYr

    def GetFiscalMnth(self,mnth):
        switcher = {
            11: 1,
            12: 2,
            1: 3,
            2: 4,
            3: 5,
            4: 6,
            5: 7,
            6: 8,
            7: 9,
            8: 10,
            9: 11,
            10: 12,
        }
        return switcher.get(mnth, "nothing")

    def GetFiscalMnth_Daily(self,mnth):
        switcher = {
            11: "November",
            12: "December",
            1: "January",
            2: "February",
            3: "March",
            4: "April",
            5: "May",
            6: "June",
            7: "July",
            8: "August",
            9: "September",
            10: "October",
        }
        return switcher.get(mnth, "nothing")

    def MailTrigger(self,mfrom,mto,subj,msg,path):
        email = mfrom
        send_to_email = mto
        subject = subj
        message = msg
        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = send_to_email
        msg['Subject'] = subject

        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition','attachment; filename="{}"'.format(os.path.basename(path)))
        msg.attach(part)

        msg.attach(MIMEText(message, 'plain'))
        server = smtplib.SMTP('smtp3.hpe.com')
        text = msg.as_string()
        server.sendmail(email, send_to_email, text)
        server.quit()

    def cleanAmount(self,amt):
        cleanAmount = amt.replace(',','')
        cleanAmount = re.sub("([0-9.]+)-$", r"-\1",cleanAmount)
        return cleanAmount

    def FetchCredentials(self,instance):
        cursor = Utility.conn.cursor()
        cursor.execute("SELECT * FROM FINANCE_CTS_RPA.dbo.Credentials where Instance = '" + instance + "'")
#        cursor.execute("SELECT * FROM FINANCE_CTS_RPA.dbo.Credentials where Instance = 'R01'")
        for row in cursor:
            cipher_suite = Fernet(Utility.key)
            pwd = (cipher_suite.decrypt(row[1].encode()))
            username = row[0]
        return username,pwd.decode()

    def CleanFolder(self,fPath):
        filesToRemove = [os.path.join(fPath,f) for f in os.listdir(fPath)]
        for f in filesToRemove:
            os.remove(f)

    def logFile(self,fname):
        logging.basicConfig( filename=fname,
                             filemode='w',
                             level=logging.DEBUG,
                             format= '%(asctime)s - %(levelname)s - %(message)s',
                           )

    def log_exception(self,e):
        logging.error("Function {function_name} raised {exception_class} ({exception_docstring}): {exception_message}".format(
                function_name = Utility.extract_function_name(self),
                exception_class = e.__class__,
                exception_docstring = e.__doc__,
                exception_message = e))

    def extract_function_name(self):
        tb = sys.exc_info()[-1]
        stk = traceback.extract_tb(tb, 1)
        fname = stk[0][3]
        return fname

    def GetFileNameFromDir(self,flpath):
        f= []
        for (dirpath, dirnames, filenames) in walk(flpath):
            f.extend(filenames)
            break

    def GetDataFromDirFiles(self,flpath, fltype, filterlist):
        df= pd.DataFrame()
        for filename in glob.glob(flpath + '\\*.' + fltype):
            if fltype == 'csv':
                print('Reading Data from file :', filename)
                df1= pd.read_csv(filename, encoding = 'ISO-8859-1')
            if fltype == 'xlsx':
                print('Reading Data from file :', filename)
                df1= pd.read_excel(filename)
            df1=df1[filterlist]
            df= pd.concat([df,df1], ignore_index= True)
        return df

    def convert_Excel_to_csv(self, filename):
        if filename[-4:]=='xlsb':
            output_filename = filename[:-4]+ 'csv'
        if filename[-3:]== 'xls':
            output_filename= filename[:-3] + 'csv'
        if filename[-4:]== 'xlsx':
            output_filename= filename[:-4] + 'csv'
        if filename[-3:]== 'csv':
            output_filename= filename
            print(filename, 'File is a csv file')
        else:
            if os.path.exists(output_filename)== False:
                import win32com.client
                excel = win32com.client.Dispatch("Excel.Application")
                excel.DisplayAlerts = False
                excel.Visible=False
                doc = excel.Workbooks.Open(filename)
                try:
                    doc.Worksheets("Sheet1").Move(Before=doc.Worksheets(1))
                except:
                    pass
            #    doc.sheets('Sheet1').Move Before: = doc.sheets(1)
                doc.SaveAs(Filename=output_filename,FileFormat=6)
                doc.Close()
                excel.Quit()
            print(filename, ' Converted to ', output_filename)
            return output_filename

    def EmptyFolder(path, allow_root=False):
        '''
        Entering "C:\\Users\\Documents\\Automagica" removes all the files and folders saved in the "Automagica" folder but maintains the folder itself.
        Standard, the safety variable allow_root is False. When False the function checks whether the path lenght has a minimum of 10 characters.
        This is to prevent entering for example "\\" as a path resulting in deleting the root and all of its subdirectories.
        To turn off this safety check, explicitly set allow_root to True. For the function to work optimal, all files present in the directory
        must be closed.
        '''
        if len(path) > 10 or allow_root:
            if os.path.isdir(path):
                for root, dirs, files in os.walk(path, topdown=False):
                    for name in files:
                        os.remove(os.path.join(root, name))
                    for name in dirs:
                        os.rmdir(os.path.join(root, name))
        return

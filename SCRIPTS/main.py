# THIS IS NOT THE MAIN CODE FILE BUT I AM MAKING THE LATEST CHANGES HERE TO CHECK 

import openpyxl
import pandas as pd
import time, subprocess, win32com.client
import warnings
from datetime import date
import os
import re
import numpy as np
import time
from config import mapping, ORDER_HDR_PATH, WORK_HDR_PATH, OP2_PATH, OP1_PATH, PA0001_PATH
from get_config import get_kwarg
import sys, os, ntpath, glob, time, subprocess, math, threading, win32com.client
import pandas as pd
from datetime import date, timedelta, datetime
from loguru import logger


warnings.filterwarnings("ignore")

configfile = "config.ini"
g = get_kwarg(configfile)
kwargs = g.get_kwargs()

instance = "DCP"
Userid = kwargs[f"{instance}_userid"]
pwd = kwargs[f"{instance}_pwd"]
date_today = date.today().strftime("%d_%m_%Y")
passing_date = datetime.today()
current_passing_date = passing_date.strftime("%m/%d/%Y")
print(Userid)

LOG_FILE_PATH = r"C:\PROJECTS\INBOUD_IDOC\LOGS\example.log"

masterDir = os.path.split(os.path.realpath(__file__))[0]

for i in range(3):
    if ntpath.basename(masterDir) == "SCRIPTS":
        masterDir = os.path.split(masterDir)[0]
        break
    else:
        masterDir = os.path.split(masterDir)[0]
print(masterDir)


InputDir = os.path.join(masterDir, "INPUT")
OutputDir = os.path.join(masterDir, "OUTPUT")


class WebsiteAutomator:
    def __init__(self):
        pass

    def saplogin(self, instance, time_out):
        try:
            path = path = kwargs["SAP_sappath"]
            subprocess.Popen(path)
        except:
            path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
            subprocess.Popen(path)
        time.sleep(time_out)

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return

        connection = application.OpenConnection(instance, True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = Userid
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = pwd

        session.findById("wnd[0]").sendVKey(0)
        time.sleep(3)
        if session.Children.Count > 1:
            if session.Children(1).Children(1).Children(0).Text == "New Password":
                session.findById("wnd[1]/usr/pwdRSYST-NCODE").Text = ""
                session.findById("wnd[1]/usr/pwdRSYST-NCOD2").Text = ""
                session.findById("wnd[1]").sendVKey(0)

            warning_msg = session.Children(1).Children(1).Children(2).Text
            print(warning_msg)
            if (
                session.Children(1).Children(1).Children(2).Text
                == "Note that multiple logons to the production system using the same user"
            ):
                session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select()
                session.findById("wnd[1]/tbar[0]/btn[0]").press()

        return session

    def download(self, session):

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "EDID4"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 5
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\IDOC_NUMBERS"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "idoc_numbers.csv"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 16
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 8
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 16
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 24
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/txtI4-LOW").text = "ZE1ORDRHDR"

        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "9999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "999999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/txtI4-LOW").setFocus()
        session.findById("wnd[0]/usr/txtI4-LOW").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\SAP"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "downloaded_data_ZE1ORDRHDR.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 20
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\IDOC_NUMBERS"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "idoc_numbers.csv"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 16
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 8
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 16
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 24
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/txtI4-LOW").text = "ZWORK_ORDER_HDR"

        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "9999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "999999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/txtI4-LOW").setFocus()
        session.findById("wnd[0]/usr/txtI4-LOW").caretPosition = 15
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\SAP"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "downloaded_data_ZWORK_ORDER_HDR.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 20
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\IDOC_NUMBERS"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "idoc_numbers.csv"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 16
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 8
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 16
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 24
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/txtI4-LOW").text = "ZE1OPERATION2"

        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "9999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "999999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/txtI4-LOW").setFocus()
        session.findById("wnd[0]/usr/txtI4-LOW").caretPosition = 15
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\SAP"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "downloaded_data_ZE1OPERATION2.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 20
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\IDOC_NUMBERS"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "idoc_numbers.csv"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 16
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 8
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 16
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 24
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/txtI4-LOW").text = "ZE1OPERATION1"

        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "9999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "999999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/txtI4-LOW").setFocus()
        session.findById("wnd[0]/usr/txtI4-LOW").caretPosition = 15
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\SAP"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "downloaded_data_ZE1OPERATION1.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 20
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass
        return session
    
    def download_psa(self, session):

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "pa0001"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 6
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\Employee_id"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "employee_id.csv"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 11
        
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/ctxtI3-LOW").text = current_passing_date
        
        session.findById("wnd[0]/usr/ctxtI3-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtI5-LOW").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[2]").press()
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell(1, "TEXT")
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "999999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\SAP"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "PA001_DATA.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass
        return session

    def download_material(self, session):

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "MARA"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 4
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\ADITIONAL_ REQUIREMENTS\material numbers"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "material_numbers.csv"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 20
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "9999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\ADITIONAL_ REQUIREMENTS\material numbers"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "downloaded_data_material_numbers.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 37
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        return session

    def download_division_data(self, session):
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZHW_PRC_SUP_PRC"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 4
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\ADITIONAL_ REQUIREMENTS\Division"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "Division.csv"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 20
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "9999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\ADITIONAL_ REQUIREMENTS\Division"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "downloaded_data_division.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 37
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        return session

    def employee_status_data(
        self, session
    ):
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "PA0000"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 4
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtI5-LOW").text = current_passing_date
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]").sendVKey(2)
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell(1, "TEXT")
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[1]/btn[2]").press()
        session.findById("wnd[1]/tbar[0]/btn[14]").press()
        session.findById("wnd[0]/usr/ctxtI5-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtI5-LOW").caretPosition = 7
        session.findById("wnd[0]/tbar[1]/btn[2]").press()
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell(1, "TEXT")
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\Employee_id"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "employee_id.csv"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 20
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()                                                                                                                                           
        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "9999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\ADITIONAL_ REQUIREMENTS\employee status data"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "downloaded_employee_status_data.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 37
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        return session

    def labour_data(self, session):
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZFILBRT"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 7
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "9999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "999999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\ADITIONAL_ REQUIREMENTS\Labour data"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZFILBRT.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZFINONLBRT"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 10
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "9999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "999999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\ADITIONAL_ REQUIREMENTS\Labour data"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZFINONLBRT.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZFICDSLBRT"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 10
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "9999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "999999999"
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\ADITIONAL_ REQUIREMENTS\Labour data"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZFICDSLBRT.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        return session

    def psa_activity_category(self, session):
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZPSA_SER_PRO_CTR"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 7
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "9999"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = "999999999"

        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\PROJECTS\INBOUD_IDOC\INPUT\PSA ACTIVITY CATEGORY"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "activity_category_data_psa.xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 31
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            button = session.findById("wnd[1]/tbar[0]/btn[11]")
            button.press()
        except:
            pass
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        return session


if __name__ == "__main__":

    logger.add(r"LOGS\file.log", rotation="00:00", retention="30 days", format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}")
    logger.info("the process has started1")
    
    automator = WebsiteAutomator()
    
    logger.info("the process has started")
    
    file_path = os.path.join('C:\\', 'PROJECTS', 'INBOUD_IDOC', 'INPUT', 'IDOC_NUMBERS', 'Book27.xlsx')
    df = pd.read_excel(file_path, index_col=None)

    file_path = os.path.join('C:\\', 'PROJECTS', 'INBOUD_IDOC', 'INPUT', 'IDOC_NUMBERS', 'idoc_numbers.csv')
    df["IDoc number"].to_csv(file_path, index=False, header=False)

    try:
        session = automator.saplogin("DCP", 20)
        
    except Exception as e:
        print("ERROR Logging to SAP: ", e)

    session = automator.download(session)
    logger.info("the process has started1")
    order_hdr_path = ORDER_HDR_PATH
    work_hdr_path = WORK_HDR_PATH
    op2_path = OP2_PATH
    op1_path = OP1_PATH
    
    
    order_hdr_df = pd.read_excel(order_hdr_path)
    psa_order_hdr = order_hdr_df[
        order_hdr_df["SDATA"].str.startswith("PSA")
    ].reset_index(
        drop=True
    )
    work_hdr_df = pd.read_excel(work_hdr_path)
    op2_df = pd.read_excel(op2_path)
    op1_df = pd.read_excel(op1_path)

    logger.info("the process has started2")
    
    OPPORTUNITY_ID = op1_df["SDATA"].str.split().str[0]
    LOG_TYPES = op1_df["SDATA"].str.split().str[1]

    contains_digit = LOG_TYPES.str.contains("\d")
    new_value = op1_df["SDATA"].str.split().str[2]
    LOG_TYPES[contains_digit] = new_value

    op1_df["opportunity_id"] = OPPORTUNITY_ID
    op1_df["log_types"] = LOG_TYPES
    op1_df.drop(columns="SDATA", inplace=True)

    logger.info("the process has started3")

    OPPORTUNITY_ID = op2_df["SDATA"].str.split().str[0]
    
    for index, row in op2_df.iterrows():
        value = row["SDATA"]
        split_values = value.split()
       
        for split_value in split_values:
            if len(split_value) == 8 and split_value.isdigit():
                employee_id = split_value
                op2_df.at[index, "employee_id"] = employee_id
            elif len(split_value) == 7 and split_value.isdigit():
                employee_id = split_value
                op2_df.at[index, "employee_id"] = employee_id
            elif len(split_value) == 6 and split_value.isdigit():
                employee_id = split_value
                op2_df.at[index, "employee_id"] = employee_id
            elif len(split_value) == 5 and split_value.isdigit():
                employee_id = split_value
                op2_df.at[index, "employee_id"] = employee_id
            else:
                pass
                
                
    op2_df["opportunity_id"] = OPPORTUNITY_ID
    op2_df.drop(columns="SDATA", inplace=True)

    op2_df = op2_df.merge(
        op1_df[["opportunity_id", "log_types"]], on="opportunity_id", how="left"
    )
    op2_df["Y-onsite , n-remining"] = op2_df["log_types"].apply(
        lambda x: "Y" if x == "Onsite" else "N"
    )

    employee_ids = op2_df['employee_id'].dropna().astype(str).str.strip()
    employee_ids = employee_ids[employee_ids != '']
    op2_df['employee_id'] = pd.to_numeric(op2_df['employee_id'], errors='coerce').astype('Int64').astype(str).replace('<NA>', '')

    employee_ids = pd.to_numeric(employee_ids, errors='coerce')

    logger.info("the process has started")

    employee_ids = employee_ids.dropna()
    employee_ids = employee_ids.astype(int)
    employee_ids = employee_ids.drop_duplicates()
    employee_ids = employee_ids.reset_index(drop=True)
    
    employee_id_csv_path = os.path.join(InputDir, "Employee_id", "employee_id.csv")
    employee_ids.to_csv(employee_id_csv_path, index=False, header=False)

    session = automator.download_psa(session)
    pa001_path = PA0001_PATH

    workbook = openpyxl.load_workbook(pa001_path)
    sheet = workbook.active

    data = sheet.values
    columns = next(data)[0:]
    data = list(data)

    pa001_df = pd.DataFrame(data, columns=columns)
        
    for index, row in pa001_df.iterrows():
        job_title_value = row['ZZ_JOBTITLE']
        if job_title_value == "TERMINATED":
            pa001_df.at[index, "ZZ_JOBLEVEL"] = "N/A"
    
    pa001_df.rename(
        columns={
            "PERNR": "employee_id",
            "PERSG": "EEGRP",
            "KOSTL": "cost center",
            "ZZ_JOBLEVEL": "Job level",
        },
        inplace=True,
    )
    
    pa001_df["employee_id"] = pa001_df["employee_id"].astype(str)
    
    
    op2_df["employee_id"] = op2_df["employee_id"].str.lstrip('0')
    op2_df = op2_df.merge(
        pa001_df[["employee_id", "EEGRP", "cost center", "Job level"]],
        on="employee_id",
        how="left",
    )

    for index, row in op2_df.iterrows():
        order_type_value = row["DOCNUM"]
        if order_type_value in psa_order_hdr["DOCNUM"].values:
            op2_df.at[index, "Order type"] = "ZPSA"
        else:
            op2_df.at[index, "Order type"] = "ZSFD"
    
    session = automator.employee_status_data(session)
    employee_status_path = os.path.join(InputDir, "ADITIONAL_ REQUIREMENTS", "employee status data", "downloaded_employee_status_data.xlsx")
    employee_status_df = pd.read_excel(employee_status_path)
    employee_status_df.rename(
        columns={"STAT2": "Employee status pa0000", "PERNR": "employee_id"},
        inplace=True,
    )
    employee_status_df["employee_id"] = employee_status_df["employee_id"].astype(str)
    op2_df = op2_df.merge(
        employee_status_df[["employee_id", "Employee status pa0000"]],
        on="employee_id",
        how="left",
    )

    op2_df["Employee status comment"] = op2_df["Employee status pa0000"].replace(mapping)
    op2_df["zfilbrt_concatenate"] = (
        op2_df["cost center"]
        + op2_df["Order type"]
        + op2_df["Job level"]
        + op2_df["Y-onsite , n-remining"]
    )
    op2_df["zfinonlbrt_concatenate"] = (
        op2_df["cost center"] + op2_df["Y-onsite , n-remining"] + op2_df["Order type"]
    )

    for index, row in op2_df.iterrows():
        employee_id_value = row["cost center"]
        if pd.isnull(employee_id_value):

            op2_df.at[index, "Employee status comment"] = (
                "EMP ID field is Blank"
            )

    material = order_hdr_df["SDATA"].str.split().str[1]
    order_hdr_df["Material"] = material
    op2_df_sfdc = op2_df[op2_df["Order type"] == "ZSFD"]
    op2_df_sfdc = op2_df_sfdc.merge(
        order_hdr_df[["DOCNUM", "Material"]], on="DOCNUM", how="left"
    )


    def clean_material(material):
        return re.sub(r'[^\x00-\x7F]+', '', material)

    filtered_cleaned_materials = [
        clean_material(material) 
        for material in op2_df_sfdc["Material"].unique() 
        if len(str(material)) <= 18
    ]

    cleaned_materials_df = pd.DataFrame(filtered_cleaned_materials, columns=["Material"])

    material_csv_path = os.path.join(InputDir, "ADITIONAL_ REQUIREMENTS", "material numbers", "material_numbers.csv")

    cleaned_materials_df.to_csv(
        material_csv_path,
        header=False, 
        index=False, 
        sep="\n" 
    )

    
    session = automator.download_material(session)
    material_path = os.path.join(
        "c:\\",
        "PROJECTS",
        "INBOUD_IDOC",
        "INPUT",
        "ADITIONAL_ REQUIREMENTS",
        "material numbers",
        "downloaded_data_material_numbers.xlsx",
    )
    material_df = pd.read_excel(material_path)
    material_df.rename(columns={"MATNR": "Material", "SPART": "Division"}, inplace=True)
    op2_df_sfdc = op2_df_sfdc.merge(
        material_df[["Material", "Division"]], on="Material", how="left"
    )
    
    division_csv_path = os.path.join(InputDir, "ADITIONAL_ REQUIREMENTS", "Division", "Division.csv")
    
    op2_df_sfdc["Division"].unique().tofile(
        division_csv_path,
        sep="\n",
        format="%s",
    )
    session = automator.download_division_data(session)

    
    division_path = os.path.join(InputDir, "ADITIONAL_ REQUIREMENTS", "Division", "downloaded_data_division.xlsx")
    division_df = pd.read_excel(division_path, engine="openpyxl")
    division_df.rename(
        columns={"ZSUPP_PRC_CTR": "profit center", "ZHW_PL": "Division"}, inplace=True
    )
    op2_df_sfdc = op2_df_sfdc.merge(
        division_df[["profit center", "Division"]], on="Division", how="left"
    )

    psa_work_order = work_hdr_df[
        work_hdr_df["DOCNUM"].isin(psa_order_hdr["DOCNUM"])
    ].reset_index()




    DEFAULT_ACTIVITY_CATEGORY_FILE_DF = pd.read_excel((os.path.join(InputDir, "Falcon S4 - PSA activity catagory list.xlsx")))
    activity_categories = DEFAULT_ACTIVITY_CATEGORY_FILE_DF["Act. Type"].tolist()
    
    def extract_activity_category(sdata):
        # Split SDATA into words and remove the last word
        words = sdata.split()[:-1]  # Removing the last word entirely

        # Check for matches using the last 2, 3, 4, or 5 words (from the remaining words)
        for num_words in range(2, 6):  # Start from 2 words to 5 words
            # If the number of words remaining is less than num_words, skip
            if len(words) < num_words:
                continue
            
            # Get the last `num_words` from the list of words (keeping the order intact)
            extracted = " ".join(words[-num_words:])
            
            # Check if this combination is in the activity categories list
            if extracted in activity_categories:
                return extracted

        # If no match is found, return None or a default value
        return "No Match"  # You can change this to any default value you prefer

    # Apply the function to each row in the psa_work_order DataFrame
    extracted_df = psa_work_order.apply(
        lambda row: pd.Series({
            "DOCNUM": row["DOCNUM"],
            "Extracted Elements": extract_activity_category(row["SDATA"])
        }),
        axis=1
    )

    print("division data has been downloaded ")

    op2_df_psa = op2_df[op2_df["Order type"] == "ZPSA"]
    op2_df_psa = op2_df_psa.merge(extracted_df[["DOCNUM", "Extracted Elements"]], on="DOCNUM", how="left")
    op2_df_psa= op2_df_psa.rename(columns={'Extracted Elements': 'Activity category'})
    
    activity_category_csv_path =  os.path.join(InputDir, "PSA ACTIVITY CATEGORY", "activity_category.csv")
    
    op2_df_psa["Activity category"].to_csv(
        activity_category_csv_path,
        header=False,
        index=False,
    )

    session = automator.psa_activity_category(session)
    
    activity_category_data_psa = os.path.join(InputDir, "PSA ACTIVITY CATEGORY", "activity_category_data_psa.xlsx")
    
    activity_category_df = pd.read_excel(
        activity_category_data_psa
    )
    activity_category_df.rename(columns={"ZPSA_ACT": "Activity category"}, inplace=True)

    op2_df_psa = op2_df_psa.merge(
        activity_category_df[["Activity category", "ZPRCTR"]],
        on="Activity category",
        how="left",
    )
    op2_df_psa.rename(columns={"ZPRCTR": "profit center"}, inplace=True)

    op2_df_sfdc["zficdslbrt_concatenate"] = (
        op2_df_sfdc["cost center"]
        + op2_df_sfdc["Y-onsite , n-remining"]
        + op2_df_sfdc["Order type"]
        + op2_df_sfdc["profit center"]
    )
    op2_df_psa["zficdslbrt_concatenate"] = (
        op2_df_psa["cost center"]
        + op2_df_psa["Y-onsite , n-remining"]
        + op2_df_psa["Order type"]
        + op2_df_psa["profit center"]
    )

    session = automator.labour_data(session)                
    
    print("labour data has been downloaded ")
    
    zfilbrt_path = os.path.join(InputDir, "ADITIONAL_ REQUIREMENTS", "Labour data", "ZFILBRT.XLSX")
    zfinonlbrt_path = os.path.join(InputDir, "ADITIONAL_ REQUIREMENTS", "Labour data", "ZFINONLBRT.XLSX")
    zficdslbrt_path = os.path.join(InputDir, "ADITIONAL_ REQUIREMENTS", "Labour data", "ZFICDSLBRT.XLSX")
    
    zfilbrt_df = pd.read_excel(zfilbrt_path)
    zfinonlbrt_df = pd.read_excel(zfinonlbrt_path)
    zficdslbrt_df = pd.read_excel(zficdslbrt_path)

    zfilbrt_df["zfilbrt_concatenate"] = (
        zfilbrt_df["KOSTL"]
        + zfilbrt_df["PSA_CLICK"]
        + zfilbrt_df["STELL"]
        + zfilbrt_df["LOG_TYPE"]
    )

    zfinonlbrt_df["zfinonlbrt_concatenate"] = (
        zfinonlbrt_df["KOSTL"] + zfinonlbrt_df["LOG_TYPE"] + zfinonlbrt_df["PSA_CLICK"]
    )

    zficdslbrt_df["zficdslbrt_concatenate"] = (
        zficdslbrt_df["KOSTL"]
        + zficdslbrt_df["LOG_TYPE"]
        + zficdslbrt_df["PSA_CLICK"]
        + zficdslbrt_df["PRCTR"]
    )

   
    op2_df_sfdc["Rate_comments"] = ""
    
    op2_df_sfdc["EEGRP"] = pd.to_numeric(op2_df_sfdc["EEGRP"], errors="coerce")
    op2_df_sfdc['EEGRP'] = op2_df_sfdc['EEGRP'].fillna(0).astype(int)
    
    for index, row in op2_df_sfdc.iterrows():
        employee_group = row["EEGRP"]
        if employee_group == 1:
            zfilbrt_concate_value = row["zfilbrt_concatenate"]
            if zfilbrt_concate_value in zfilbrt_df["zfilbrt_concatenate"].values:
                op2_df_sfdc.at[index, "Rate_comments"] = "the cost/rate is present in zfilbrt"
                
            if zfilbrt_concate_value not in zfilbrt_df["zfilbrt_concatenate"].values:
                op2_df_sfdc.at[index, "Rate_comments"] = "the cost/rate is not present in zfilbrt"                
                
        elif employee_group == 2:
            zfinonlbrt_concate_value = row["zfinonlbrt_concatenate"]
            if zfinonlbrt_concate_value in zfinonlbrt_df["zfinonlbrt_concatenate"].values:
                op2_df_sfdc.at[index, "Rate_comments"] = "the cost/rate is present in zfinonlbrt"
                
            if zfinonlbrt_concate_value not in zfinonlbrt_df["zfinonlbrt_concatenate"].values:
                op2_df_sfdc.at[index, "Rate_comments"] = "the cost/rate is not present in zfinonlbrt"                
                
        elif employee_group == 3 or employee_group == 4:
            zficdslbrt_concatenate_value = row["zficdslbrt_concatenate"]
            if zficdslbrt_concatenate_value in zficdslbrt_df["zficdslbrt_concatenate"].values:
                op2_df_sfdc.at[index, "Rate_comments"] = "the cost/rate is present in zficdslbrt"
                
            if zficdslbrt_concatenate_value not in zficdslbrt_df["zficdslbrt_concatenate"].values:
                op2_df_sfdc.at[index, "Rate_comments"] = "the cost/rate is not present in zficdslbrt"                
    
    op2_df_psa["Rate_comments"] = ""
    
    op2_df_psa["EEGRP"] = pd.to_numeric(op2_df_psa["EEGRP"], errors="coerce")
    op2_df_psa['EEGRP'] = op2_df_psa['EEGRP'].fillna(0).astype(int)
                
    for index, row in op2_df_psa.iterrows():
        employee_group = row["EEGRP"]
        if employee_group == 1:
            zfilbrt_concate_value = row["zfilbrt_concatenate"]
            
            if zfilbrt_concate_value in zfilbrt_df["zfilbrt_concatenate"].values:
                op2_df_psa.at[index, "Rate_comments"] = "the cost/rate is present in zfilbrt"            
            

            if zfilbrt_concate_value not in zfilbrt_df["zfilbrt_concatenate"].values:
                op2_df_psa.at[index, "Rate_comments"] = "the cost/rate is not present in zfilbrt"
                
        elif employee_group == 2:
            zfinonlbrt_concate_value = row["zfinonlbrt_concatenate"]
 
            if zfinonlbrt_concate_value in zfinonlbrt_df["zfinonlbrt_concatenate"].values:
                op2_df_psa.at[index, "Rate_comments"] = "the cost/rate is present in zfinonlbrt"           
            
            
            if zfinonlbrt_concate_value not in zfinonlbrt_df["zfinonlbrt_concatenate"].values:
                op2_df_psa.at[index, "Rate_comments"] = "the cost/rate is not present in zfinonlbrt"
                
        elif employee_group == 3 or employee_group == 4:
            zficdslbrt_concatenate_value = row["zficdslbrt_concatenate"]
            
            if zficdslbrt_concatenate_value  in zficdslbrt_df["zficdslbrt_concatenate"].values:
                op2_df_psa.at[index, "Rate_comments"] = "the cost/rate is present in zficdslbrt"

            if zficdslbrt_concatenate_value not in zficdslbrt_df["zficdslbrt_concatenate"].values:
                op2_df_psa.at[index, "Rate_comments"] = "the cost/rate is not present in zficdslbrt"


    final_output_path = os.path.join(OutputDir, "output.xlsx")
    with pd.ExcelWriter(r"OUTPUT\FINAL_OUTPUT\output.xlsx") as writer:
        op2_df_sfdc.to_excel(writer, sheet_name="SFDC", index=False)
        op2_df_psa.to_excel(writer, sheet_name="PSA", index=False)
        
    print("############################END OF THE CODE ######################################")
    
    
    
    
    
    
    



import warnings
warnings.filterwarnings("ignore")
import sys, os, ntpath, glob, logging, time, subprocess, win32com.client, math
logging.warn(os.path.realpath(__file__))
import traceback

masterDir= os.path.split(os.path.realpath(__file__))[0]
#masterDir= os.path.split(os.curdir)[0]
for i in range(3):
    if ntpath.basename(masterDir)=="SCRIPTS":
        masterDir= os.path.split(masterDir)[0]
        break
    else:
        masterDir= os.path.split(masterDir)[0]
        
logging.warn(masterDir)    
inputDir = os.path.join(masterDir, "INPUT")
scriptsDir = os.path.join(masterDir, "SCRIPTS")
print(masterDir)

print(scriptsDir)
outputDir = os.path.join(masterDir, "OUTPUT")
log_dir = os.path.join(masterDir, "Log_File")

sys.path.append(scriptsDir)

#%% Setting Log Decorator and Log Files
# import log_decorator
# from log import log
# log= log()

#%%
'''Importing Miscellenious file'''
from MISC import Utility
u= Utility()

'''class function for config'''
class get_kwarg(object):
    def __init__(self,configfile):
        import configparser
        self.config= configparser.ConfigParser()
        self.config.read(configfile)
        logging.warn(ntpath.abspath(configfile))

        self.config_master=os.path.join(scriptsDir,"config.ini")
        self.config1 = configparser.ConfigParser()
        self.config1.read(self.config_master)
        
    '''Getting the master config'''
    def get_kwarg1(self):
        try:

            kwargs = dict(self.config.items('DEFAULT'))
            for each_section in self.config.sections():
                dict1= dict(list(set(self.config.items(each_section))-set(self.config.items('DEFAULT'))))
                dict1 = {each_section + "_" + key:value for (key,value) in dict1.items()}
                exec("%s=%s" % (each_section,dict1))
                exec(("kwargs.update(%s)" % (each_section)))
            return kwargs
        except Exception:
                excp = "Exception occurs due to absence of Master configfile   - " + traceback.format_exc()
                logging.warn(excp)
    
    '''Adding Master config with main config file'''
    def get_kwargs(self):
        try:
            kwargs= self.get_kwarg1()
            print('"helooo')
            for each_section in self.config1.sections():
                dict1= dict(list(set(self.config1.items(each_section))-set(self.config1.items('DEFAULT'))))
                dict1 = {each_section + "_" + key:value for (key,value) in dict1.items()}
                exec("%s=%s" % (each_section,dict1))
                exec(("kwargs.update(%s)" % (each_section)))
            # logging.warn("config file read. Total keys in config : ",len(kwargs))
            # logging.warn(kwargs)
            return kwargs   
    
        except Exception:
                excp = "Exception occurs due to not finding Main config file and Master configfile   - " + traceback.format_exc()
                logging.warn(excp)
    
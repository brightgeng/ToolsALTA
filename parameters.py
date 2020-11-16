import datetime
import os
import sys
import time

DATE0 = datetime.datetime(1970, 1, 1)
DIR = r"c:\Users\runda\OneDrive - RundaTech\04 工作\0405 艾格威贸易\ALTA_Matching"
YMD = time.strftime("%Y%m%d", time.localtime())
DOUT = sys.stdout
DB = os.path.join(DIR, 'log', 'AltaTools.db')

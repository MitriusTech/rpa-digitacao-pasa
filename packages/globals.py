from datetime import datetime, timedelta
import os
from timeit import default_timer as timer

# Globals
today_ = datetime.today()
tomorow = (today_ + timedelta(days=1)).strftime('%d/%m/%Y')
today = today_.strftime('%Y%m%d')
todayFormatted = today_.strftime('%d/%m/%Y')
path = os.getcwd()
temp = f'{path}\\temp'
log = f'{path}\\log\\{today}-{today_.strftime("%H%M")}' 
logFile = f'log{today}{today_.strftime("%H%M%S")}.log'
logFileFullPath = f'{log}\\{logFile}'
SAPSession = None
Wnd0 = None
Menubar = None
UserArea = None
Statusbar = None
UserName = None
startTime = timer()  
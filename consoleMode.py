from autoRptService import * 
import win32api, sys

logger.addScreenMode()
logger.setDebug()



class debugMode(autoRptServ):
    def __init__(self):
        self.stop_requested = False 
    

if __name__ == '__main__':
    def on_exit(sig, func=None):
        print ('exit signal received, exit...')
        sys.exit()
    win32api.SetConsoleCtrlHandler(on_exit, True)
    sv = debugMode()
    sv.main()

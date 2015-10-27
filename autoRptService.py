import win32serviceutil
import win32service
import win32event
import servicemanager
import time, os, datetime
import task, atlogger
from threading import Thread
from config import config





def taskproc(job):
    t = task.task(job, logger)
    t.work()

class autoRptServ(win32serviceutil.ServiceFramework):
    _svc_name_ = 'AutoReport Service'
    _svc_display_name_ = 'AutoReport Service'
    _logger = atlogger.g_logger
    
    def __init__(self, args ):
        win32serviceutil.ServiceFramework.__init__(self, args )
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.stop_requested = False

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        self.stop_requested = True

    def SvcDoRun(self):
        servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_,'')
                )
        self.main()


    def main(self):
        if not hasattr(self, 'stop_event'):
            self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self._logger.info('*** Autoreport Service Start ***')
        c = config()
        if c.loadCfg() is False: return
        self.cfg = c
        while True:
            now = datetime.datetime.now()
            jobs = self.cfg.tasks(now)
            if jobs is None or len(jobs) == 0:
                self._logger.info('No task added at all, exit')
                return
            secs = jobs[0][1]
            self._logger.info('Wait ' + str(secs) + ' seconds to do next task')
            val = win32event.WaitForSingleObject( self.stop_event, 
                    secs*1000 )
                    #1000)
            if val == win32event.WAIT_OBJECT_0 or val == win32event.WAIT_ABANDONED:
                self._logger.info(' %%% Autoreport Service Stopped %%%')
                break
            for i in jobs:
                self._logger.info('Wake up and smell the coffee, a thread is working on a task')
                p = Thread(target = taskproc, args = (i[0],))
                p.daemon = True
                p.start()
            time.sleep(1)


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(autoRptServ)

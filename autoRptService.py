import win32serviceutil
import win32service
import win32event
import servicemanager
import time, os, json, math, enum, datetime, calendar
import task
from atlogger import atLogger
from dateutil.relativedelta import relativedelta
from threading import Thread

logger = atLogger()

class WORKTYPE(enum.Enum):
    DAILY = 1
    WEEKLY = 2
    MONTHLY = 3
    YEARLY = 4

class workTime:
    def __init__(self, timestr):
        self._logger = logger
        self.month, self.day = 0, 0
        s = timestr.strip().split(',')
        if len(s) == 0: 
            raise Exception('Serious error, must stop', 'in autoRptService.py')
        if s[0].strip() == 'daily':
            self.worktype = WORKTYPE.DAILY
            self.getTime(s[1])
        elif s[0].strip() == 'weekly':
            self.worktype = WORKTYPE.WEEKLY
            self.day = int(s[1].strip())
            self.getTime(s[2])
        elif s[0].strip() == 'monthly':
            self.worktype = WORKTYPE.MONTHLY
            self.day = int(s[1].strip())
            self.getTime(s[2])
        elif s[0].strip() == 'yearly':
            self.worktype = WORKTYPE.YEARLY
            ss = s[1].strip()
            if len(ss) != 4: 
                raise Exception('Serious error, month and date must be wrong', 'in autoRptService.py')

            self.month = int(ss[:2])
            self.day = int(ss[2:])
            self.getTime(s[2])
        else:
            raise Exception('Unrecognised time format', 'in autoRptService.py')
        self._logger.info('Task time initialised')
        self._logger.info('Type: ' + str(self.worktype) + ' mon/day: ' + str(self.month) + '/' + str(self.day) + ' hours: ' + str(self.hours) + ' mins:' + str(self.mins) + ' secs:' + str(self.seconds))
    
    def getTime(self, timestr):
        s = timestr.strip()
        if len(s) != 8: raise Exception('Time format error','')
        ss = s.split(':')
        self.hours = int(ss[0])
        self.mins = int(ss[1])
        self.seconds = int(ss[2])
        self.secOfDay = self.hours*3600 + self.mins*60 + self.seconds

    def upcoming(self, cur):
        now, timeMiss, diff = cur, False, None
        if now.hour*3600 + now.minute*60 + now.second > self.secOfDay: timeMiss = True
        dst = datetime.datetime(now.year,now.month,now.day,self.hours,self.mins,self.seconds)
        if self.worktype == WORKTYPE.DAILY:
            if timeMiss: 
                dst = dst + datetime.timedelta(days=1)
            diff = dst - cur
        elif self.worktype == WORKTYPE.WEEKLY:
             n = (self.day - 1 - now.weekday())%7
             if n == 0: 
                 if timeMiss: 
                     dst = dst + datetime.timedelta(days=7)
             else:
                 dst = dst + datetime.timedelta(days=n)
             diff = dst - cur
        elif self.worktype == WORKTYPE.MONTHLY:
            day = self.day
            a = calendar.monthrange(now.year, now.month)
            if a[1] < self.day: day = a[1]
            if day == now.day:
                if timeMiss:
                    dst = dst + relativedelta(months=1)
                diff = dst - cur
            elif day > now.day:
                dst = dst + datetime.timedelta(days=(day-now.day))
                diff = dst - cur
            else:
                tmp = datetime.datetime(now.year,now.month,1,self.hours,self.mins,self.seconds)
                tmp = tmp + relativedelta(months=1)
                if tmp.day == self.day:
                    diff = tmp - cur
                else:
                    a = calendar.monthrange(tmp.year, tmp.month)
                    if a[1] < self.day: day = a[1]
                    else: day = self.day
                    tmp = tmp + datetime.timedelta(days=(day-tmp.day))
                    diff = tmp - cur
        elif self.worktype == WORKTYPE.YEARLY:
            day = self.day
            a= calendar.monthrange(now.year, self.month)
            if a[1] < self.day: day = a[1]
            dateMiss = datetime.date(now.year,now.month,now.day) - datetime.date(now.year, self.month, day)
            if dateMiss.total_seconds() == 0:
                if timeMiss:
                    a= calendar.monthrange(now.year+1, self.month)
                    if a[1] < self.day: day = a[1]
                    dst = datetime.datetime(now.year+1, self.month, day, self.hours, self.mins, self.seconds)
            elif dateMiss.total_seconds() < 0:
                dst = datetime.datetime(now.year, self.month, day, self.hours, self.mins, self.seconds)
            else:
                a= calendar.monthrange(now.year+1, self.month)
                if a[1] < self.day: day = a[1]
                dst = datetime.datetime(now.year+1, self.month, day, self.hours, self.mins, self.seconds)
            diff = dst - cur
        return diff



class config:
    def __init__(self):
        self._logger = logger
        # create needed dirs if not exist
        sqldir = os.path.dirname(os.path.realpath(__file__))  + os.sep + "sql"
        exceldir = os.path.dirname(os.path.realpath(__file__))  + os.sep + "excel"
        if not os.path.exists(sqldir):
            os.mkdir(sqldir)
        if not os.path.exists(exceldir):
            os.mkdir(exceldir)

    def prepare(self, task):
        self.prepareCfgItem(task,'sender')
        self.prepareCfgItem(task,'mailpass')
        self.prepareCfgItem(task,'mailtitle')
        self.prepareCfgItem(task,'mailtext')
        self.prepareCfgItem(task,'mailserver')
        self.prepareCfgItem(task,'receiver')
        self.prepareCfgItem(task,'cc')
        self.prepareCfgItem(task,'bcc')
        self.prepareCfgItem(task,'dbconn')

    def prepareCfgItem(self, task, item):
        it = task.get(item, self.cfg.get(item))
        if it is not None: task[item] = it


    def tasks(self, tm):
        jobs, closest = None, 1000000000
        for t in self.cfg.get('tasks'):
            if t.get('enabled','true').lower() != 'true': continue
            w = t.get('period')
            if w is None: continue
            diff = w.upcoming(tm) 
            if diff.total_seconds() < closest:
                jobs = [(t,diff.total_seconds())] 
                closest = diff.total_seconds()
            elif diff.total_seconds() == closest:
                jobs.insert(len(jobs), (t, diff.total_seconds()))
        return jobs

    def loadCfg(self):
        with open(os.path.dirname(os.path.realpath(__file__)) + os.sep + 'config.json') as f:
            try:
                self.cfg = json.load(f)
                for t in self.cfg['tasks']:
                    if t.get('period') is None: raise Exception('No time specified','config.json')
                    w = workTime(t.get('period'))
                    t['period'] = w
                    self.prepare(t)
            except Exception as e:
                self._logger.exception('Something wrong when loading config file')
                self._logger.exception(e)
                return False
        return True



def taskproc(job):
    t = task.task(job, logger)
    t.work()

class autoRptServ(win32serviceutil.ServiceFramework):
    _svc_name_ = 'AutoReport Service'
    _svc_display_name_ = 'AutoReport Service'
    _logger = logger
    
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

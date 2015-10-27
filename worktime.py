import enum, datetime, calendar, time, math, atlogger
from dateutil.relativedelta import relativedelta
class WORKTYPE(enum.Enum):
    DAILY = 1
    WEEKLY = 2
    MONTHLY = 3
    YEARLY = 4

class workTime:
    def __init__(self, timestr):
        self._logger = atlogger.g_logger
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
        if now.hour*3600 + now.minute*60 + now.second + now.microsecond/1000000.0 > self.secOfDay: timeMiss = True
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


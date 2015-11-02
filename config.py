from worktime import workTime
import os, atlogger, json
class config:
    def __init__(self):
        self._logger = atlogger.g_logger
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

    def tasksAll(self):
        for t in self.cfg.get('tasks'):
            if t.get('enabled','true').lower() != 'test': continue
            w = t.get('period')
            if w is None: continue
            yield t


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



import sendmail
import pyodbc
import os, datetime
from openpyxl import Workbook

class task:
    def __init__(self, t, logger):
        self.taskinfo = t
        self.logger = logger

    def resultIter(self, cursor, size=300):
        while True:
            results = cursor.fetchmany(size)
            if not results: break
            for result in results:
                yield result


    def sendFile(self, fname):
        login = self.taskinfo.get('sender'), self.taskinfo.get('mailpass')
        cc = self.taskinfo.get('cc')
        bcc = self.taskinfo.get('bcc')
        to = self.taskinfo.get('receiver')
        sender = self.taskinfo.get('sender')
        subject = self.taskinfo.get('mailtitle')
        text = self.taskinfo.get('mailtext')
        server = self.taskinfo.get('mailserver')
        sendmail.sendmail(login, sender, to, cc, bcc, subject, text, fname, server)  

    def work(self):
        try:
            sql = None
            curdir = os.path.dirname(os.path.realpath(__file__)) 
            con_str = self.taskinfo.get('dbconn')
            con = pyodbc.connect(con_str) 
            cursor = con.cursor()
            wb = Workbook(optimized_write = True)
            for sheet in self.taskinfo.get('sheets'):
                sqlfile = os.path.join(curdir,'sql') + os.sep + sheet.get('sql')
                with open(sqlfile) as fn:
                    sql = fn.read()
                ws = wb.create_sheet()
                ws.title = sheet.get('name')
                cursor.execute(sql)
                columns = [column[0] for column in cursor.description]
                # Make column header
                ws.append(columns)
                for r in cursor:
                    ws.append(list(r))

            con.close()
            # Generate excel file
            fname = os.path.join(curdir,'excel') + os.sep + self.taskinfo.get('excel') + '-' + datetime.datetime.now().strftime('%Y%m%d-%H%M%S.xlsx')
            wb.save(filename = fname )
            self.sendFile([fname])
        except Exception as e:
            self.logger.exception(e)

if __name__ == '__main__':
    pass



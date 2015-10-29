import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

def sendmail( login, sendfrom, sendto, cc, bcc, subject, text, files=None, server="smtp.gmail.com:587"):
    assert isinstance(sendto, list)
    msg = MIMEMultipart(
            From = sendfrom, 
            To = COMMASPACE.join(sendto),
            Date = formatdate(localtime = True),
            )
    msg['Subject'] = subject
    msg['From'] = sendfrom
    msg['To'] = ', '.join(sendto)
    if cc: 
        msg.add_header('Cc', ', '.join(cc))
    msg.attach(MIMEText(text))

    sendto = sendto + cc + bcc

    for f in files or []:
        with open(f, 'rb') as fil:
            msg.attach(MIMEApplication(
                fil.read(),
                Content_Disposition='attachment; filename="%s"' % basename(f),
                Name=basename(f)
                ))

    smtp = smtplib.SMTP(server)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(login[0], login[1])
    smtp.sendmail(sendfrom, sendto, msg.as_string())
    smtp.close()


if __name__ == "__main__":
    pass

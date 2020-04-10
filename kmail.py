#! python3
"""
Send emails to list of recepients
Create a log email of activities and mail to the user

"""
import datetime as dt
import smtplib

class mail():
    """ userinfo is a dict of user information : 'name' and 'pass' req\n
        emails is a list of string email addresses\n
        content is a dict of 'subject' : string, and 'body' : string
    """
    def __init__(self, userinfo: dict, emails: list, content: dict):
        self.user = userinfo['name']
        self.password = userinfo['pass']
        self.smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
        self.__login_email__()
        self.emails = emails
        self.subject = 'Subject: ' + content['subject'] + '\n'
        self.content = self.subject + content['body']
        self.statusLog = []
        self.send_emails()
        self.email_status()
        self.smtpObj.quit()

    def __login_email__(self):
        """ Logs user into mail server """
        print('Connecting to email server... ', end='', flush=True)
        self.smtpObj.ehlo()
        self.smtpObj.starttls()
        self.smtpObj.login(self.user, self.password)
        print('Done')
    
    def send_emails(self):
        """ Sends data to list of emails """
        for email in self.emails:
            body = ''
            for line in self.content:
                body += line
            message = f'Sending email to {email}... '
            print(message, flush=True, end='')
            #sendmailStatus = smtpObj.sendmail(user, email, body) # turn on to send emails
            sendmailStatus = {} # turn off when sending emails
            if sendmailStatus != {}:
                status = 'Incomplete'
                message = message + status
                self.statusLog.append(message)
                print(status)
                message = f'There was a problem sending email to {email}: {sendmailStatus}'
                self.statusLog.append(message)
                print(message)
            else:
                status = 'Completed'
                message = message + status
                self.statusLog.append(message)
                print(status)
        self.statusLog.append('\n\n' + body)
        return

    def email_status(self):
        """ Send completed status email to user containing details from statusLog """
        print(f'Sending status email to {self.user}... ', end='', flush=True)
        now = dt.datetime.now().strftime('%I:%M:%S%p on %x')
        body = f'{self.subject} -- STATUS EMAIL\n'
        body = body + f'\n{self.subject} emails sent at {now}\n\n'
        for line in self.statusLog:
            body += line + '\n'
        body += '\nEND OF STATUS'
        sendmailStatus = self.smtpObj.sendmail(self.user, self.user, body)
        if sendmailStatus != {}:
            print('Incomplete')
            print(f'There was a problem sending the status email to {self.user}: {sendmailStatus}')
        else:
            print('Completed')
        return

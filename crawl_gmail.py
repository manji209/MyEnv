

import smtplib
import time
import imaplib
import email
import credentials
from email.parser import HeaderParser

import base64
import os


# -------------------------------------------------
#
# Utility to read email from Gmail Using Python
#
# ------------------------------------------------


def read_email_from_gmail():
    ORG_EMAIL = "@lalucky.com"
    FROM_EMAIL = "dinh.ho" + ORG_EMAIL
    FROM_PWD = "He11sh0ck209!@#"
    SMTP_SERVER = "imap.gmail.com"
    SMTP_PORT = 993

    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox')

        type, data = mail.search(None, 'ALL')
        #mail_ids = data[0]

        for num in data[0].split():
            typ, data = mail.fetch(num, '(RFC822)')
            print('Message %s\n%s\n' % (num, data[0][1]))
        mail.close()
        mail.logout()


        '''
        for num in data[0].split():
            typ, data = mail.fetch(num, '(RFC822)')

            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1])
                    email_subject = msg['subject']
                    email_from = msg['from']
                    print('From : ' + email_from + '\n')
                    print('Subject : ' + email_subject + '\n')
  



        id_list = mail_ids.split()
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])


        for i in range(latest_email_id,latest_email_id - 5, -1):
            typ, data = mail.fetch(str(i), '(RFC822)')

            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(str(response_part[1]))
                    email_subject = msg['subject']
                    email_from = msg['from']
                    print(email_subject)
                    print(email_from)
                    #print('From : ' + email_from + '\n')
                    #print('Subject : ' + email_subject+ '\n')
'''

    except Exception as e:
        raise
        #print(str(e))


read_email_from_gmail()

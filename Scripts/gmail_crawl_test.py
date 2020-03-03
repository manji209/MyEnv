import smtplib
import time
import imaplib
import email
import re
import datefinder
import pandas as pd


# -------------------------------------------------
#
# Utility to read email from Gmail Using Python
#
# ------------------------------------------------

def read_email_from_gmail():
    ORG_EMAIL = "@lalucky.com"
    FROM_EMAIL = "kenny.nguyen" + ORG_EMAIL
    FROM_PWD = "KNtkeo51*"
    SMTP_SERVER = "imap.gmail.com"
    SMTP_PORT = 993
    results = []
    headers = ['Ref #', 'ETD', 'ETA']

    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox')

        type, data = mail.search(None, '(SINCE "01-JAN-2019")')
        mail_ids = data[0]

        id_list = mail_ids.split()


        for i in reversed(id_list):
            typ, data = mail.fetch(i, '(RFC822)' )
            temp_result = []


            for response_part in data:
                if isinstance(response_part, tuple):

                    #line to decode from byte to string
                    try:
                        msg = email.message_from_string(response_part[1].decode('utf-8'))

                        email_subject = msg['subject']

                        if email_subject and "ETD" in email_subject:
                            # Extract Reference number from subject
                            ref_num = re.search("[^\s]+", email_subject)
                            temp_result.append(ref_num.group()[1:])
                            print(temp_result[0])

                            sub_string = email_subject.partition("]")[2]

                            matches = datefinder.find_dates(sub_string)

                            for match in matches:
                                print(match.date().strftime("%m/%d"))
                                temp_result.append(match.date().strftime("%m/%d"))

                            # email_from = msg['from']
                            # print ('From : ' + email_from + '\n')
                            print('Subject : ' + email_subject + '\n')

                    except UnicodeDecodeError:
                        continue


            if len(temp_result) == 3:
                results.append(temp_result)


    except Exception as e:

        raise

    df = pd.DataFrame(results, columns=headers)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('../Out/EMAIL.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()

read_email_from_gmail()
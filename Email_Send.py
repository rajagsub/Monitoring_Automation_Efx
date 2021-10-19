"""
    Email_Send.py, by Rajagopalan S(AF40411), Thirumavalavan N (AG13731), 2021-04-20.

    This program sends the EMAIL notification for EFX Error & Consolidated Reporting.

    Pre-requisite : Need to import below libraries with python ver 3.8.8

    Input : Uses the Consolidated_Trigger.txt file with the list of the current time, Execution indicator

"""

import smtplib
from email.mime.multipart import MIMEMultipart
import email.mime.application
from email.mime.text import MIMEText
from datetime import datetime
import pandas as pd

def trigger_email():
    msg = MIMEMultipart()
    msubject = "EFX Gateway File Routing Notification_{}".format(datetime.now().strftime('%Y%m%d-%H%M%S'))

    # setup the parameters of the message
    msg['From'] = "AF40411@ANTHEM.COM"
    msg['To'] = "AG18144@ANTHEM.COM"
    msg['Cc'] = "AF40411@ANTHEM.COM"
    msg['Subject'] = msubject

    # Email Body
    email_html = open('Email_Body1.html')
    message = email_html.read()

    # add in the message body
    msg.attach(MIMEText(message, 'html'))
    toaddrs = msg['Cc'].split(",") + [msg['To']]


    for filename in attachments:
        f = filename
        fo = open(f, 'rb')
        attach = email.mime.application.MIMEApplication(fo.read(), _subtype="xlsx")
        fo.close()
        attach.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(attach)

    print("Sending Email...")

    # create server
    try:
        server = smtplib.SMTP('smtpinternal.wellpoint.com', 25)
        server.sendmail(msg['From'], toaddrs, msg.as_string())
        print ("Successfully sent email")
        del msg

    except Exception as e:
        print(e)
        print("Error: unable to send email")

    finally:
        print ("Closing the server...")
        server.quit()


def check_complete_trigger_email(trigger_time, trigger_exec):

    if current_time >= trigger_time and trigger_exec == 'N':
        trigger_exec = 'Y'
        update_file_content = trigger_time + ',' + trigger_exec
        print(update_file_content)
        consolidated_trigger_file = open("Consolidated_Trigger.txt", "w")
        consolidated_trigger_file.write(update_file_content)
        consolidated_trigger_file.close()

        # Appending the consolidated execl sheet to the ATTACHMENTS List.
        attachments.append("EFX_Consolidated_Report.xlsx")

    if trigger_exec == 'Y' and current_time < trigger_time:
        # Preparing for the next run.
        trigger_exec = 'N'
        update_file_content = trigger_time + ',' + trigger_exec
        consolidated_trigger_file = open("Consolidated_Trigger.txt", "w")
        consolidated_trigger_file.write(update_file_content)
        consolidated_trigger_file.close()


if __name__ == '__main__':
    print('--- Starting Emailing Program ---')
    now = datetime.now()  # current date and time
    current_time = now.strftime("%H:%M:%S")
    print("Current_time:", current_time)

    attachments = []
    Error_EFX_df = pd.read_excel('EFX_Routing_and_Failed_Files_Report.xlsx')

    # Reading the Trigger file for the consolidation report.
    consolidated_trigger_file = open("Consolidated_Trigger.txt", "r")
    string_list = consolidated_trigger_file.readlines()
    for i in string_list:
        temp = i.split(",")
        trigger_time = temp[0]
        trigger_exec = temp[1]
    consolidated_trigger_file.close()


    if Error_EFX_df.empty:
        print("File is empty")
        # Check if we need to attach Consolidated report
        check_complete_trigger_email(trigger_time, trigger_exec)

        if len(attachments) > 0:
            trigger_email()

    else:
        attachments = ['EFX_Routing_and_Failed_Files_Report.xlsx']

        # Check if we need to attach Consolidated report
        check_complete_trigger_email(trigger_time, trigger_exec)

        # Trigger email
        trigger_email()
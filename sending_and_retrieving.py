#pip install --upgrade pywin32
import os
import win32com.client as win32
import win32com.client
import datetime
from pathlib import Path

def saveattachemnts(messages,today,path):
    for message in messages.Items:
        if message.Subject == 'FORD_FCSD – In-Transit report '+ str(today):
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment)))
                if message.Unread:
                    message.Unread = False
                break
def main():
    #receive email
    path = r"C:\\Users\\yoeri.samwel\\Documents\\task_scheduler_tasks\\FORD_FCSD_in_transit_shipment_report_error_check\\data"
    today = datetime.date.today()
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    saveattachemnts(inbox,today,path)

    #processing input


    #email out
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0  # size of the new email
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Testing Mail'
    newmail.To = 'yoeri.samwel@dsv.com'
    newmail.CC = 'yoeri.samwel@dsv.com'
    newmail.Body = 'Hello, this is a test email to showcase how to send emails from Python and Outlook.'
    # attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    # newmail.Attachments.Add(attach)
    filename = 'FORD_FCSD - In-Transit Shipment Report 2023-09-12_2.xlsx'
    attach = os.path.join('C:\\Users\\yoeri.samwel\\Documents\\task_scheduler_tasks\\FORD_FCSD_in_transit_shipment_report_error_check', filename)
    newmail.Attachments.Add(attach)
    newmail.Display()
    newmail.Send()

if __name__=="__main__":
    main()
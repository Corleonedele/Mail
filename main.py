import win32com.client
import os
import sys

department = "安保部"
number = 0

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
accounts = mapi.Folders

namelist = open("information of applicant.txt", "w")
package = os.makedirs("attachment")
path = os.getcwd()+"\\attachment"

def check_department(title):
    global department
    if department in str(title):
        return True
    else: 
        return False

def check_attachments(mail):
    global number
    try:
        attachments = mail.Attachments
        number = len([x for x in attachments])
        return True
    except:
        return False

def write_info(mail):
    namelist.write(mail.SenderName)
    namelist.write(str(mail))
    namelist.write("\n")

def download_atta(mail):
    global number
    for x in range(1, number + 1):
        attachment = mail.Attachments.Item(x)
        attachment.SaveASFile(os.path.join(path, attachment.FileName))

inbox = accounts[1].Folders
mails = inbox[1].Items

for mail in mails:
    if check_department(mail):
        write_info(mail)
        print(mail)
        if check_attachments(mail):
            download_atta(mail)

print("Finish Successfully")




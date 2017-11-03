#specify path that contains .msg files
#this code will create a .cv file that contains
#Does not work with arabic

import win32com.client
import re, os

filename = "Sheet.cv"
file = open(filename, 'w')

f = os.listdir()#get name of all .msg files in path
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

tmp = 'Date#To#CC#Subject#Body\n'
file.write(tmp)

for i in f:
    if not i.endswith('.msg'):
        continue
    try:
        print(i)
        msg = outlook.OpenSharedItem(i)
        tmp = str(msg.SentOn) + '#' + msg.To + '#' + msg.CC + '#' + msg.Subject + '#' + msg.Body
        tmp = tmp.replace('\n', '')
        tmp = tmp.replace('\r', '')
        tmp = re.split("NADER QASSEMACCOUNTS MANAGER", tmp.upper())
        tmp  = re.split("GROUP ACCOUNTING", tmp[0].upper())    
        tmp = tmp[0] + "\n"
        file.write(tmp)
        del msg
    except Exception as e:
        print(e)
        file.close()
    break


file.close()
del outlook

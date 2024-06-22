import win32com.client as win32
import os

outlook = win32.Dispatch('outlook.application')
namespace = outlook.GetNamespace('MAPI')
directory = "G:\\GitHub\\JIRA-creator\\sample folder"

accounts = namespace.Accounts


for i in range(accounts.Count):
    print(f'Account {i + 1}: {accounts.Item(i + 1).DisplayName}')

account_index = 1  

for fileName in os.listdir(directory):
    if fileName.lower().endswith('.pdf'):
        print(fileName)
        mail = outlook.CreateItem(0)
        mail.Subject = fileName
        mail.Attachments.Add(directory + "\\" + fileName)
        # mail.Body = 'This is a test email sent from a Python script using Outlook!'
        mail.To = 'gurditrajat13@gmail.com'
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, accounts.Item(account_index)))


        try:
            mail.Send()
            print('Email sent successfully!')
        except Exception as e:
            print(f'Error: {e}')

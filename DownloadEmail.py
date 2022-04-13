from email import message
import datetime
import win32com.client as client
import pandas as pd

today = datetime.date.today()

# create instance of Outlook
outlook = client.Dispatch('Outlook.Application')

# get the inbox
namespace = outlook.GetNameSpace('MAPI')
inbox = namespace.GetDefaultFolder(6)


# the email I want to download a file from

# get only mail items from the inbox (other items can exists and will return an error if you try get the subject line of a non-mail item)
mail_items = [item for item in inbox.Items if item.Class == 43]


# filter to the target email
filtered = [item for item in mail_items if item.Unread and item.Senton.date() == today]

print(filtered)

if len(filtered) == 0:
        print ("No Attachment")
n=0
# get the first item if it exists (assuming the there is only one item to get)
while n < len(filtered):

    if len(filtered) != 0:
        target_email = filtered[n]
        n+=1
    
# get attachments

        #if target_email.Attachments.Count > 0:
        attachments = pd.read_html(target_email.HTMLBody)   

        #print(attachments) 
    
    # save attachments to file
        # save_path = r'C:\Users\Yusuf_Budiawan\Documents\Factory-Work-Plan-Consolidate\Factory-Work-Plan-Consolidate\downloaded items\{}'


        # for file in attachments:
        #         file.SaveAsFile(save_path.format(file.FileName))

    elif len(filtered) != 0:
        print ("No Attachment")
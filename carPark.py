from pathlib import Path    #core python module
import win32com.client      #pip install win32

def getEmails():
    #output directory
    output_dir = Path.cwd() / "Output"
    output_dir.mkdir(parents=True, exist_ok=True)

    #connect to outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    #connect to folder
    inbox = outlook.GetDefaultFolder(6)

    #get messages
    messages = inbox.Items

    for message in messages:
        subject = message.Subject
        sender_email = message.SenderEmailAddress
        attachments = message.Attachments
        sent_time = message.SentOn

        if (sender_email != "eastlansing@harveyec.com" and \
            "Customizable Data Export" not in subject):
            continue
        else:
            title = subject[-7:] + " " + str(sent_time)

        # if sent_time < last_ran:
        

        #create seperate target folders
        target_folder = output_dir / str(subject)
        target_folder.mkdir(parents=True, exist_ok=True) 

        #save attachemnts
        for attachment in attachments:
            attachment.SaveAsFile(target_folder / str(attachment))

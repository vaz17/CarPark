from pathlib import Path    #core python module
import win32com.client      #pip install win32

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
    body = message.Body
    attachments = message.Attachments

    #create seperate target folders
    target_folder = output_dir / str(subject)
    target_folder.mkdir(parents=True, exist_ok=True) 

    #write body to text file
    Path(target_folder / "EMAIL_BODY.txt").write_text(str(body))

    #save attachemnts
    for attachment in attachments:
        attachment.SaveAsFile(target_folder / str(attachment))

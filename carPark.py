from pathlib import Path
import win32com.client
import pandas as pd

# Initialize Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

# Path to the existing CSV file
existing_csv_path = Path("path/to/existing_file.csv")

# Loop through the messages in the inbox
for message in messages:
    subject = message.Subject
    sender_email = message.SenderEmailAddress
    sent_time = message.SentOn  # This is a datetime object

    # Filter emails based on the sender and subject
    if (sender_email != "eastlansing@harveyec.com" and
        "Customizable Data Export" not in subject):
        continue

    # Get the attachment
    attachments = message.Attachments
    if attachments.Count == 1:
        attachment = attachments.Item(1)
        
        # Save the attachment to a temporary location
        temp_file_path = Path.cwd() / attachment.FileName
        attachment.SaveAsFile(str(temp_file_path))

        # Load the attachment CSV into a pandas DataFrame
        attachment_df = pd.read_csv(temp_file_path)

        # Add the sent_time column to the DataFrame
        attachment_df['sent_time'] = str(sent_time)

        # Append to the existing CSV
        if existing_csv_path.exists():
            existing_df = pd.read_csv(existing_csv_path)
            combined_df = pd.concat([existing_df, attachment_df], ignore_index=True)
        else:
            combined_df = attachment_df

        # Save the updated CSV file
        combined_df.to_csv(existing_csv_path, index=False)

        # Optionally, remove the temporary file
        temp_file_path.unlink()  # Remove the temporary file
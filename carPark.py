from pathlib import Path
import win32com.client
import pandas as pd
from collections import Counter
from datetime import datetime, timezone

# Function to read the last run time from lastRan.txt
def get_last_run_time(file_path):
    try:
        with open(file_path, 'r') as file:
            last_run_str = file.read().strip()  # Read and strip any whitespace
            # Convert string to datetime object, assuming the format is 'YYYY-MM-DD HH:MM:SS'
            last_run_time = datetime.strptime(last_run_str, '%Y-%m-%d %H:%M:%S')
            return last_run_time
    except FileNotFoundError:
        # If the file doesn't exist, assume this is the first run
        return None

def countPark(file_path):
    with open(file_path, 'r') as file:
        lines = file.readlines()

    # Skip the header and create a list to store car park names
    car_park_names = []

    # Extract car park names from each line
    for line in lines[1:]:  # Skip the header
        name, _ = line.split(';')
        car_park_names.append(name.strip().strip('"'))

    # Count occurrences of each car park name
    car_park_count = Counter(car_park_names)

    # Convert the Counter to a DataFrame
    car_park_df = pd.DataFrame(car_park_count.items(), columns=['car_park_name', 'count'])

    # Print the results
    return car_park_df


# Initialize Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

# Path to the existing CSV file
existing_csv_path = Path("./TEST.csv")

# Path to the last run file
last_run_file_path = Path("./lastRan.txt")

# Get the last run time
last_run_time = get_last_run_time(last_run_file_path)

# Loop through the messages in the inbox
for message in messages:
    subject = message.Subject
    sender_email = message.SenderEmailAddress
    sent_time = message.SentOn  # This is a datetime object


    # Filter emails based on the sender and subject
    if (sender_email != "eastlansing@harveyec.com" and
        "Customizable Data Export" not in subject):
        continue

    # Convert sent_time to UTC just to be sure, even though it's likely already in UTC
    sent_time = sent_time.astimezone(timezone.utc)

    # Ensure last_run_time is also in UTC
    last_run_time = last_run_time.astimezone(timezone.utc)

    # Now they can be compared safely
    if last_run_time and sent_time < last_run_time:
        continue  # Skip this email if it was sent before or at the last run time

    # Check if the email was sent after the last run
    if last_run_time and sent_time < last_run_time:
        continue  # Skip this email if it was sent before or at the last run time


    # Extract the date from sent_time
    date_only = sent_time.strftime('%Y-%m-%d')  # Format to 'YYYY-MM-DD'

    # Get the hour from the last two characters of the subject
    hour = subject[-2:]  # Assuming the last two characters are the hour

    # Get the attachment
    attachments = message.Attachments
    if attachments.Count == 1:
        attachment = attachments.Item(1)
        
        # Save the attachment to a temporary location
        temp_file_path = Path.cwd() / attachment.FileName
        attachment.SaveAsFile(str(temp_file_path))

        # Load the attachment file into a pandas DataFrame
        attachment_df = countPark(temp_file_path)

        # Add the date and hour columns to the DataFrame
        attachment_df['sent_date'] = date_only
        attachment_df['sent_hour'] = hour

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

# Write the current time as the last run time to lastRan.txt
with open(last_run_file_path, 'w') as file:
    file.write(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))a
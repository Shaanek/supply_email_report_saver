# %%
#######################################################################################
# OUTLOOK EMAIL ATTACHMENT EXTRACTOR
# This script helps to download attachments from Outlook emails automatically
# Prerequisites: 
# 1. Install pywin32 library using: pip install pypiwin32
# 2. Must have Microsoft Outlook installed
#######################################################################################

import win32com.client  # Library to interact with Microsoft Office applications
import datetime
import os

#######################################################################################
# STEP 1: Get Today's Date
#######################################################################################
# Format today's date as DD-MM-YYYY (e.g., 01-12-2023)
dateToday = datetime.datetime.today()
FormatedDate = (
    '{:02d}'.format(dateToday.day) + '-' + 
    '{:02d}'.format(dateToday.month) + '-' + 
    '{:04d}'.format(dateToday.year)
)

#######################################################################################
# STEP 2: Connect to Outlook
#######################################################################################
# Create connection to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the root folder (email account)
# Note: Item(1) is usually the default account
# Change the number if you want to access a different email account
# You can print root_folder to see the email address and verify
root_folder = outlook.Folders.Item(2)  # Change 2 to 1 if this is your primary account

# Print all available folders in the email account
print("Available folders in your account:")
for folder in root_folder.Folders:
    print(folder.Name)

#######################################################################################
# STEP 3: Access Specific Folder
#######################################################################################
# Replace 'Your_Folder_Name' with the actual folder name where your emails are stored
# This could be 'Inbox', 'Archive', or any custom folder
target_folder_name = 'Your_Folder_Name'
inbox = root_folder.Folders[target_folder_name]

# Get all messages and sort them by received time (newest first)
messages = inbox.Items
messages.Sort("[ReceivedTime]", False)

#######################################################################################
# STEP 4: Function to Save Attachments
#######################################################################################
def save_attachments(subject, file_name):
    """
    Save attachments from emails matching the specified subject.
    
    Parameters:
    subject (str): The exact subject line of the email to search for
    file_name (str): The name to save the attachment as
    """
    
    # Create output directory if it doesn't exist
    output_dir = os.path.join(os.path.expanduser("~"), "Downloads", "Outlook_Attachments")
    os.makedirs(output_dir, exist_ok=True)
    
    found = False
    # Search through all messages
    for message in messages:
        if message.Subject == subject:
            found = True
            print(f"Found email with subject: {subject}")
            
            # Check if email has attachments
            if message.Attachments.Count > 0:
                # Save each attachment
                for attachment in message.Attachments:
                    save_path = os.path.join(output_dir, file_name)
                    attachment.SaveAsFile(save_path)
                    print(f"Saved attachment as: {save_path}")
                    break  # Remove this if you want to save all attachments
            else:
                print("No attachments found in the email")
            break  # Remove this if you want to process all matching emails
    
    if not found:
        print(f"No email found with subject: {subject}")

#######################################################################################
# STEP 5: Execute the Script
#######################################################################################
# Print the email account we're working with
print(f"Working with email account: {root_folder}")

# Example usage:
# Parameters:
# 1. Email subject to search for
# 2. Name to save the file as
save_attachments(
    subject="Your Email Subject Here",  # Replace with the exact email subject
    file_name="Downloaded_Attachment.xlsx"  # Replace with desired filename
)

#######################################################################################
# USAGE INSTRUCTIONS:
# 1. Install required library: pip install pypiwin32
# 2. Modify the following variables according to your needs:
#    - root_folder = outlook.Folders.Item(2) -> Change number if needed
#    - target_folder_name = 'Your_Folder_Name' -> Change to your folder name
#    - In save_attachments() call:
#      * Replace "Your Email Subject Here" with the actual email subject
#      * Replace "Downloaded_Attachment.xlsx" with desired filename
# 3. Run the script
#######################################################################################
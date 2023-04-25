import os
import re
from pathlib import Path
import win32com.client # "pip install pywin32" to use this package

# define directories to save attachments
user_downloads_folder = os.path.join(Path.home(), "Downloads")
output_directory = os.path.join(user_downloads_folder, "Outlook Attachments")
if not os.path.exists(output_directory):
    os.makedirs(output_directory)
    
# dictionary of default folders in Outlook
# taken from https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
outlook_folder_dictionary = {
    "Calendar": 9,
    "Conflicts": 19,
    "Contacts": 10,
    "DeletedItems": 3,
    "Drafts": 16,
    "Inbox": 6,
    "Journal": 11,
    "Junk": 23,
    "LocalFailures": 21,
    "ManagedEmail": 29,
    "Notes": 12,
    "Outbox": 4,
    "SentMail": 5,
    "ServerFailures": 22,
    "SuggestedContacts": 30,
    "SyncIssues": 20,
    "Tasks": 13,
    "ToDo": 28,
    "FoldersAllPublicFolders": 18,
    "RssFeeds": 25,
}


def main():
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Connect to Inbox folder
    inbox_folder = outlook.GetDefaultFolder(outlook_folder_dictionary["Inbox"])
    
    # Connect to subfolder inside of Inbox
    # Change dummy to the actual folder name
    try:
        target_folder = inbox_folder.Folders["dummy"]
    except Exception as e:
        print("Error opening folder. Make sure name is correct.")
        exit()
    
    # Get messages in target folder
    messages = target_folder.Items
    
    # Loop through messages. Save attachments if criteria met.
    for message in messages:
        subject = message.Subject
        body = message.body
        attachments = message.Attachments
        sent_date = message.SentOn
        sent_date_formatted = re.sub("[^0-9]+", "_", str(sent_date))
        unread = message.UnRead
        sender = message.Sender
                
        if not unread:
            # skip messages already read
            continue
        
        if not attachments:
            # skip messages with no attachments
            continue
        
        # # example of other things you can do
        # if not "search_term" in body:
        #     continue
        
        for attachment in attachments:
            # remove any special characters from attachment name
            attachment_filename = re.sub("[^0-9a-zA-Z\.]+", "", attachment.FileName)
            
            # add date to attachment name
            new_attachment_filename = "_".join([sent_date_formatted, attachment_filename])

            # set output folder
            attachment_output_folder = os.path.join(output_directory, "dummy")
            if not os.path.exists(attachment_output_folder):
                os.makedirs(attachment_output_folder)
                
            # save attachment
            attachment_output_filepath = os.path.join(attachment_output_folder, new_attachment_filename)                    
            attachment.SaveAsFile(attachment_output_filepath)
            
if __name__ == "__main__":
    main()
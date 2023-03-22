# OutlookEmail Python Code - get the list of attachments of an email
This Python code uses the win32com.client library to interact with Microsoft Outlook and retrieve the attachments from an email with a specified subject in a specified folder.

### Prerequisites
- Python 3.x
- win32com package

### Usage
Open a command prompt or terminal window.
Navigate to the directory containing the Python script.
Run the following command to install the win32com package: pip install pywin32
Run the following command to execute the script: python outlook_email.py
Enter the name of the Outlook folder when prompted.
Enter the subject of the email when prompted.
The script will then print out the titles of any attachments in the specified email.
Note: This script currently only works on Windows machines with Microsoft Outlook installed and configured.

### Explanation of Code
The OutlookEmail class represents an email in a specified folder in Outlook.
The __init__ method of this class takes in the outlook object, the name of the folder, and the subject of the email.
The __init__ method then retrieves the specified folder from the outlook object, retrieves all messages in that folder, finds the last message with the specified subject, and retrieves its attachments.
The _find_subfolder method recursively searches for a subfolder with the specified name.
The _find_message_by_object_text method searches through all messages in the specified folder for a message with the specified subject.
The print_attachments method prints out the titles of all attachments in the specified email.
The script prompts the user for the folder name and email subject, creates an instance of the OutlookEmail class with these values, and then calls the print_attachments method on this instance.



import win32com.client as win32

class OutlookEmail:
    def __init__(self, outlook, folder_name, object_text):
        self.folder = self._find_subfolder(outlook.Folders, folder_name)
        self.messages = self.folder.Items
        self.last_message = self._find_message_by_object_text(self.messages, object_text)
        self.attachments = self.last_message.Attachments
    
    def _find_subfolder(self, folders_obj, folder_search_name):
        for i in range(0, len(folders_obj)):
            try:
                ret = folders_obj[i].Folders[folder_search_name]
                return ret
            except:
                ret = self._find_subfolder(folders_obj[i].Folders, folder_search_name)
            if ret is not None:
                return ret
            else:
                continue
    
    def _find_message_by_object_text(self, messages, object_text):
        for message in messages:
            if message.Subject == object_text:
                return message
        raise Exception(f"No email found with subject '{object_text}' in folder '{self.folder.Name}'")

    def print_attachments(self):
        attachment_titles = "Vous trouverez ci-joint la liste des documents: \n"
        for attachment in self.attachments:
            attachment_titles += attachment.DisplayName + "\n"
        print(attachment_titles)

# outlook object
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# prompt user for folder name and email subject
folder_name = input("Enter the name of the Outlook folder: ")
object_text = input("Enter the subject of the email: ")

# create OutlookEmail instance for the specified folder and email
email = OutlookEmail(outlook, folder_name, object_text)

# print attachment titles for the specified email
email.print_attachments()

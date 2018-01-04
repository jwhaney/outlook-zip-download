'''
This script uses the win32com python module to find the weekly email in Outlook containing the zip shapefile needed.
It is downloaded to a specified directory. This script can be used for other types of email attachments.

author: john haney
'''

import win32com.client

def main():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    root_folder = namespace.Folders.Item(1)
    #specify subfolder names if you want, otherwise just use the Inbox and remove the dot notation subfolders
    subfolder = root_folder.Folders['Inbox'].Folders['Subfolder'].Folders['Sufolder']

    messages = subfolder.Items
    #specifiy the directory you want the zip to be downloaded to
    dir = 'C:\\Insert_Your\\Directory\\Here\\'

    latestEmail = messages.GetLast()

    for attachment in latestEmail.Attachments:
        attachment.SaveAsFile(dir + 'Your_File_Name.zip') #specify the file name
        print 'download from outlook successful'

main()

import win32com.client
from win32com import *
from win32api import *
from win32com.client import *
import os



outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #gets the folder namespace for the current outlook instance
folder = outlook.Folders["YOUR BASE FOLDER"]  #gets the folder for the given string. This represents the BASE folder you want archived -- all subfolder within will be archived as well
downloadpath = "C:\\Users\\YOUR.NAME\\Downloads\\FILEPATH" #Path where you want to archive emails and associated email file structure

def generatePath(folderName):
    filepath = downloadpath + folderName
    filepathClean = ''.join(c for c in filepath if c not in '<>') #removes characters that can't be used in filenames in Windows
    os.makedirs(filepathClean)

def downloadEmails(mailitems, folderName):
    mailitemTotal = mailitems.Items.Count
    for mailitemcount in range(0, mailitems.Items.Count): #for range loop 
        if mailitemcount < 100000: #will stop downloading after the 100 thousandth email within a given mailbox subfolder
            mailitemsubject = None
            try:
                filepath = downloadpath + folderName
                mailitemDate = mailitems.items[mailitemcount].CreationTime
                mailitemDateAppend = str(mailitemDate.month) + str(mailitemDate.year)
                mailitemsubject = mailitems.items[mailitemcount].Subject #gets the subject of email as a string
                mailitemsubjectClean = ''.join(c for c in mailitemsubject if c not in '/\:*?"<>|') #removes characters that can't be used in filenames in Windows
                filepathsave = filepath + "\\" + mailitemDateAppend + mailitemsubjectClean + ".msg" 
                if len(filepathsave) > 250:    
                    filepathsavetruncated = filepathsave[:250] + ".msg" 
                else:
                    filepathsavetruncated = filepathsave
                mailitems.items[mailitemcount].SaveAs(filepathsavetruncated) #saves the email as the subject name 
                print(str(mailitemcount) + " of " + str(mailitemTotal) + " saved " + " in " + str(mailitems))
            except:
                pass
        else:
            return

def folderloop(folderobject, recursionCount=0):
    initialFolderList = folderobject
    #identView = "---"
    for foldercount in range(0, initialFolderList.Count):
        if initialFolderList.Count > 0:
            generatePath(str(initialFolderList[foldercount].FolderPath)) #generates file path
            downloadEmails(initialFolderList[foldercount], str(initialFolderList[foldercount].FolderPath)) #saves files into folder path
            folderloop(initialFolderList[foldercount].Folders, recursionCount+1) #recursive function which repeats the folderloop for every nested folder
        else:
            print(str(initialFolderList[foldercount])) #this else handles folders which do not have nested folders, and the initial loop should continue
            generatePath(str(initialFolderList[foldercount].FolderPath))
            downloadEmails(initialFolderList[foldercount], str(initialFolderList[foldercount]))
    
folderloop(folder.Folders) #calls the folderloop function with the folder.Folders object


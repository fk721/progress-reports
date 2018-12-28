from __future__ import print_function
import smtplib
from email.mime.text import MIMEText
import os
import io
from apiclient.http import MediaIoBaseDownload
from googleapiclient.discovery import build
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client import file, client, tools
from httplib2 import Http
import xlrd
import datetime
import unicodedata
class Date:
    def __init__(self, month, day, year):
        self.month = int(month)
        self.day = int(day)
        self.year = int(year)
    def __lt__(self, rhs):
        if (self.year < rhs.year):
            return True
        elif (self.year > rhs.year):
            return False
        if (self.month < rhs.month):
            return True
        elif (self.month > rhs.month):
            return True
        if (self.day < rhs.day):
            return True
        else:
            return False
    def __eq__(self, rhs):
        return (self.month == rhs.month and self.day == rhs.day and self.year == rhs.year)
    def __str__(self):
        return (str(self.month) + "/" + str(self.day) +"/" +str(self.year))
    def isValid(self):
        return (1 <= self.month <= 12)
    def __repr__(self):
        return str(self)
def clear():
    os.system( 'cls' )
def downloadFiles():
    print("Downloading all of the report cards listed under KTRH Progress Reports")
    folderId = '1LlFeBANo-xVrfi7woI9kgu0FZ97F-rSu' # Unique to KTRH
    SCOPES = 'https://www.googleapis.com/auth/drive.readonly'
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('drive', 'v3', http=creds.authorize(Http()))
    children = service.files().list(q = "'1LlFeBANo-xVrfi7woI9kgu0FZ97F-rSu' in parents").execute()
    if ('files' not in children):
        print("No report cards in this directory")
        return
    collectionID = children['files']
    arrID = []
    arrNames = []
    for item in collectionID:
        arrID.append(item['id'])
        arrNames.append(item['name'].strip() + ".xlsx")
    for i in range(len(arrID)):
        file_id = arrID[i]
        request = service.files().export_media(fileId=file_id,mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        fh = io.FileIO(arrNames[i], 'wb')
        downloader = MediaIoBaseDownload(fh , request)
        done = False
        print("Downloading",arrNames[i])
        while done is False:
            status, done = downloader.next_chunk()
    clear()
    print("Finished downloading",len(collectionID),"report cards from KTRH Progress Reports")
def checker():
    month,date,year = input("Enter the desired date in MM/DD/YY format: ").split("/")
    userDate = Date(month, date, year)
    if (not userDate.isValid()):
        print("Not a valid date!")
        return False
    elif (userDate.year == 18):
        if (userDate.month == 12):
            colOrg = 4
            endCol = 14
        else:
            print("Sorry! I cannot search there!")
            return False
    elif (userDate.year == 19):
        colOrg = userDate.month * 12 + 4
        endCol = colOrg + 10
    else:
        print("Sorry! I cannot search there!")
        return False
    print("Checking for students who scored below an 85% or less on", str(userDate))
    scoreSheet = open("gradeFile.txt", "w")
    for filename in os.listdir(os.getcwd()):
        if filename.endswith(".xlsx"):
            print("Checking file:",filename)
            workbook = xlrd.open_workbook(filename)
            reportCard = workbook.sheet_by_index(0)
            studentName = str(reportCard.cell_value(0,0)).strip()
            done = True
            col = colOrg
            while (done):
                if (col >= endCol):
                    done = False
                elif (reportCard.cell_value(col,0)): 
                    date = reportCard.cell_value(col,0)
                    year, month, day, hour, minute, second = xlrd.xldate_as_tuple(date, workbook.datemode)
                    excelDate = Date(month,day,int(year) - 2000)
                    if (excelDate > userDate):
                        done = False
                    elif(excelDate == userDate):
                        if isinstance(reportCard.cell_value(col,4), str):
                            cwValue = [0,"N/A"]
                        else:
                            cwValue = [(reportCard.cell_value(col,4)) * 100]
                        if isinstance(reportCard.cell_value(col,5), str):
                            hwValue = [0,"N/A"]
                        else:
                            hwValue = [(reportCard.cell_value(col,5)) * 100]
                        if (cwValue[0] <= 85 or hwValue[0] <= 85):
                            scoreSheet.write(studentName.upper() + ": " + str(excelDate) +
                                             " Classwork grade: " + str(cwValue[0]) + "% "
                                             + "Homework grade: " + str(hwValue[0]) + "% " + "\n")
                col += 1
    scoreSheet.close()
    clear()
    print("Done searching through all of the report cards")
    return True
def deleteFiles():
    print("Removing report cards from the OS")
    cwd = os.getcwd()
    count = 0
    for item in os.listdir(cwd):
        if item.endswith('.xlsx'):
            print("Removing",item)
            count +=1
            os.remove(os.path.join(cwd, item))
    clear()
    print("Removed",count,"files from the OS")
    clear()
    return count

def main():
    downloadFiles()
    res = checker()
    count = deleteFiles()
    if (res):
        print("Successfully searched through",count,"report cards and removed them from the operating system.")
        print("Enter 'gradeFile.txt' on the command line to view the progress reports that require attention.")
main() 

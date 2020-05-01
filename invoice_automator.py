from __future__ import print_function
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient import discovery
import pickle
import os.path
from apiclient.http import MediaFileUpload
from apiclient.http import MediaIoBaseDownload
import io
import time
from argparse import ArgumentParser
from gooey import Gooey
import sys
import re


class ProgressBar(object):
    DEFAULT = 'Progress: %(bar)s %(percent)3d%%'
    FULL = '%(bar)s %(current)d/%(total)d (%(percent)3d%%) %(remaining)d to go'

    def __init__(self, total, width=40, fmt=DEFAULT, symbol='=',
                 output=sys.stderr):
        assert len(symbol) == 1

        self.total = total
        self.width = width
        self.symbol = symbol
        self.output = output
        self.fmt = re.sub(r'(?P<name>%\(.+?\))d',
            r'\g<name>%dd' % len(str(total)), fmt)

        self.current = 0

    def __call__(self):
        percent = self.current / float(self.total)
        size = int(self.width * percent)
        remaining = self.total - self.current
        bar = '[' + self.symbol * size + ' ' * (self.width - size) + ']'

        args = {
            'total': self.total,
            'bar': bar,
            'current': self.current,
            'percent': percent * 100,
            'remaining': remaining
        }
        print('\r' + self.fmt % args, file=self.output)

    def done(self):
        self.current = self.total
        self()
        print('', file=self.output)


'''
(1)Create a new folder named NEWMONTH 'YEAR, with all invoices
     and receipts from the previous month

(2)First find the folder ID of the newmonth folder
   Then grab all of the files in the new month folder
   Change the name of each of the new files according to the month, 
       based on the file being a receipt or invoice
   For each name that is changed, save it to a filenamearray

(3)For each name in the filenamearray, change the necesssary fields
   using the sheets api

(4)Use the drive API to convert all files into PDFS of the same name 
'''
@Gooey(program_name= "revolv x bali Operation Form Automator",  header_height = 100, default_size = [600, 500] )
def parse_args():
    parser = ArgumentParser(description="Now is the time to update the 'Master Invoices and Receipts' Folder."+"\nPlease ensure that each spreadsheet is filled out correctly \nand is formatted '##_Inv/Rec_Master_StoreName'.")
    parser.add_argument('Parent_Folder_Name', help="Please enter the name of the foler you wish to create (eg 'Month 20XX')")
    user_inputs = vars(parser.parse_args())
    return str(user_inputs['Parent_Folder_Name'])

if __name__ == '__main__':
    parentFolderName = parse_args()
    
    driveCreds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            driveCreds = pickle.load(token)
        
    drive_service = discovery.build('drive', 'v3', credentials= driveCreds)
    sheetScope = ['https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('balidatacreds.json', sheetScope)
    
    months = ['January', 'February', 'March', 'April', 
              'May', 'June', 'July', 'August', 'September',
              'October', 'November', 'December']
    
    
    
    raw_store_names = []
    page_token = None
    query = "'1HIVHpNmOQAw5Dx8HQ4CbUb34bqNDgUWA' in parents and (trashed = false) and (mimeType = 'application/vnd.google-apps.spreadsheet') "
    while True:
        response = drive_service.files().list(q = query , 
                                              spaces = 'drive', 
                                              fields = 'nextPageToken, files(name)',
                                              pageToken = page_token).execute()
        
        for thing in response.get('files', []):
            tempName = thing.get('name')
            raw_store_names.append(tempName)
            if len(tempName.split("_")) < 2:
                print('Found Faulty Store Name on: '+tempName)
                break
        
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break
        
    store_name_and_index = [] 
    number_of_stores = 0
    for name in raw_store_names:   
        split_name = name.split('_')
        store_name_and_index.append([int(split_name[0]), split_name[-1]])
        if int(split_name[0]) > number_of_stores:
            number_of_stores = int(split_name[0])
      
    stores = [None] * number_of_stores
    for name in store_name_and_index:
        stores[name[0]-1] = name[1]
    
    print('')
    print("The folder '"+parentFolderName+"' will be created in 'Operations/Bali/Restaurants Invoice and Receipts/'" )
    
    year = parentFolderName[-2:] # the 'XX' in the name above
    monthNumber = ''
    
    for index, month in enumerate(months):
        if month in parentFolderName:
            monthNumber = str(index+2)
            
    #monthNumber = '09' ## eg '08' for July
    dayNumber = '5' 
    monthName = months[int(monthNumber)-1]
    print('')
    
    invoiceFileNames = []
    receiptFileNames = []
    
    for index, store in enumerate(stores):
        invoiceFileNames.append('Invoice_#'+str(index+1)+"-SI"+"-"+monthNumber+"-"+year+'_'+store)
        receiptFileNames.append('Receipt_#'+str(index+1)+'-SR'+"-"+monthNumber+"-"+year+'_'+store)
        #who cares about consistency with "" and ''
    
    print("Here's what the computer has found:")
    for thing in invoiceFileNames:
        print(thing)
    for thing in receiptFileNames:
        print(thing)
    print('')
    
    #given a string which contains a store name, find the store number corresponding to store name
    #uses stores array above
    def store_number(storeName):
        trailingName = storeName[-5:]
        for index, store in enumerate(stores):
            if trailingName in store:
                return index + 1
            
        return None
    
    #creates the new file name based on the master receipt or invoice of a given store
    #uses receiptFileNames and invoiceFileNames arrays above
    def new_file_name(tempFileName):
        trailingName = tempFileName.split('_')[-1].lower()
        
        if 'Rec' in tempFileName:
            for rname in receiptFileNames:
                lower_rname = rname.lower()
                if trailingName in lower_rname:
                    return rname
            
        if 'Inv' in tempFileName:
            for iname in invoiceFileNames:
                lower_iname = iname.lower()
                if trailingName in lower_iname:
                    return iname
        
        print('returned NoneType on' + tempFileName)
        return None
    
    #'Master Invoices and Receipts' id as of July 16th: 1HIVHpNmOQAw5Dx8HQ4CbUb34bqNDgUWA
    #'Restaurant Invoice and Receipt' id as of July 17th: 1V7wIo1KZTSCIZeUCxfZ6YCOrggECIptU
    #'testing' id as of August 6th: 1cd3dTDcEvXetVIQoYSOj7_yGX7ZQUdkI
    #(1) Copying the master sheet
    
    print("Creating a copy of the 'Master Invoices and Receipts' folder:")
    masterChildren = []
    page_token = None
    query = "'1HIVHpNmOQAw5Dx8HQ4CbUb34bqNDgUWA' in parents and (trashed = false) and (mimeType = 'application/vnd.google-apps.spreadsheet') "
    while True:
        response = drive_service.files().list(q = query , 
                                              spaces = 'drive', 
                                              fields = 'nextPageToken, files(*)',
                                              pageToken = page_token).execute()
        
        for thing in response.get('files', []):
            masterChildren.append(thing)
        
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break
    
    file_metadata = {
         'name' : parentFolderName,
         'mimeType' : 'application/vnd.google-apps.folder',
         'parents' : ['1V7wIo1KZTSCIZeUCxfZ6YCOrggECIptU'] 
    } #Note that 'parents' is the Restaurants Invoice and Receipt folder id
    
    newMonthFile = drive_service.files().create(body = file_metadata, fields='id').execute()
    
    progress = ProgressBar(len(masterChildren), fmt=ProgressBar.FULL)
    
    for child in masterChildren:
        progress.current += 1
        progress()
        
        file_metadata = {
            'name' : child.get('name'),
            'parents' : [newMonthFile.get('id')]
        }
        
        drive_service.files().copy(fileId = child.get('id'), body = file_metadata).execute()
        time.sleep(100/number_of_stores)
    print('')
    
    
    #(2) Finding Parent Folder ID
    query = "name = '"+ parentFolderName + "'and (trashed = false) and ('1V7wIo1KZTSCIZeUCxfZ6YCOrggECIptU' in parents)"
    parentFolderID = None
    page_token = None
    while True:
        response = drive_service.files().list(q = str(query), 
                                              spaces = 'drive', 
                                              fields = 'nextPageToken, files(id, name)',
                                              pageToken = page_token).execute()
        
        for thing in response.get('files', []):
            parentFolderID = thing.get('id')
            print('')
            print('Found folder: ' + parentFolderName)
            break
        
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break
    print('')
    
    #Finding all files within the Parent Folder ID, names and Ids
    childrenNames = []
    childrenIds = []
    query = "'"+parentFolderID+"' in parents and (trashed = false) and (mimeType = 'application/vnd.google-apps.spreadsheet') "
    
    while True:
        response = drive_service.files().list(q = query , 
                                              spaces = 'drive', 
                                              fields = 'nextPageToken, files(id, name)',
                                              pageToken = page_token).execute()
        
        for thing in response.get('files', []):
            tempName = thing.get('name')
            childrenNames.append(tempName)
            childrenIds.append(thing.get('id'))
        
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break
    print('')
    sheetsToEdit = []
    progress = ProgressBar(len(masterChildren), fmt=ProgressBar.FULL)
    #Updating the names for the receipts and invoices:
    print("Preparing each file in "+ parentFolderName+":")
    
    
    for index, file_id in enumerate(childrenIds):
        progress.current += 1
        progress()
        sheetName = new_file_name(childrenNames[index])
        tempFile = drive_service.files().get(fileId=file_id).execute()
        temp_file = drive_service.files().update(fileId = file_id, body = {'name' : sheetName}).execute()
        sheetsToEdit.append(sheetName)
    print('')
    print('')
    
    #(3) Editing each sheet!
    
    progress = ProgressBar(len(masterChildren), fmt=ProgressBar.FULL)
    print('Editing each invoice and receipt:')
    
    for thing in childrenNames:
        print(thing)
    
    erroredSheets = []
    for sheet in sheetsToEdit:
        
        progress.current += 1
        progress()
        
        invoiceDate = monthNumber+"/"+dayNumber+"/"+year
        dueDate = monthNumber+"/"+str(int(dayNumber)+14)+"/"+year
        
        try: 
            currentWorksheet = gspread.authorize(creds).open(sheet).sheet1 
            #if the sheet is an invoice 
            if 'Inv' in sheet[0:3]:
                #'Invoice #' in the form of 'storeNumber/SI/monthNumber/Year'
                #'Invoice Date' in the form of 'monthNumber/dayNumber/Year'
                #'Due Date' in the form of '7/19/2019'
                
                invoiceNumber = str(store_number(sheet))+"/SI/"+monthNumber+"/"+year
                
                dateCell = currentWorksheet.find('Invoice Date:')
                currentWorksheet.update_cell(dateCell.row, dateCell.col+1, invoiceDate)
                
                invoiceNoCell = currentWorksheet.find('Invoice #:')
                currentWorksheet.update_cell(invoiceNoCell.row, invoiceNoCell.col+1, invoiceNumber)
                
                dueDateCell = currentWorksheet.find('Due Date:')
                currentWorksheet.update_cell(dueDateCell.row, dueDateCell.col+1, dueDate)
                
            #if the sheet is a receipt
            if 'Rec' in sheet[0:3]:
                #'Receipt Number' is 'Invoice Number'
                #'Date' is 'Due Date'
                #'For Payments' is "Pembayaran Invoice" + 'Invoice #'
                invoiceNumber = str(store_number(sheet))+"/SI/"+monthNumber+"/"+year
                
                receiptNoCell = currentWorksheet.find('Receipt Number')
                currentWorksheet.update_cell(receiptNoCell.row,  receiptNoCell.col+2, invoiceNumber)
                
                dateCell = currentWorksheet.find('Date')
                currentWorksheet.update_cell(dateCell.row, dateCell.col+2, dueDate)
                
                paymentCell = currentWorksheet.find('For Payments')
                currentWorksheet.update_cell(paymentCell.row, paymentCell.col+2, "Pembayaran Invoice "+invoiceNumber)
            time.sleep(100/number_of_stores)
        except:
            erroredSheets.append(sheet)
            
        
    for thing in erroredSheets:
        print(thing)
        
    print('')
    print('')
    print('Exporting and Downloading each sheet as PDF:')
    progress = ProgressBar(len(masterChildren), fmt=ProgressBar.FULL)
    
    erroredSheets = []
    for index, childrenId in enumerate(childrenIds):
        progress.current += 1
        progress()
        try:
            file_id = childrenId
            request = drive_service.files().export_media(fileId = file_id, mimeType = 'application/pdf')
            fh = io.FileIO('pdfs/'+sheetsToEdit[index], 'wb')
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
        except:
            erroredSheets.append(childrenNames[index])
     
    print("AN ERROR MIGHT HAVE OCCURED ON THE FOLLOWING SHEETS")
    print("CHECK THEM BY HAND ON GOOGLE DRIVE PLEASE:")
    print(erroredSheets)
    
    print('')
    print('')
    print("Uploading PDFS of Receipts and Invoices to Google Drive")
    progress = ProgressBar(len(masterChildren), fmt=ProgressBar.FULL)
    
    erroredSheets = []
    for sheet in sheetsToEdit:
        progress.current += 1
        progress()
        try:
            file_metadata = {'name': sheet, 'parents' : [parentFolderID]}
            media = MediaFileUpload('pdfs/'+sheet, mimetype = 'application/pdf')
            
            uploadFile = drive_service.files().create(body = file_metadata, media_body = media).execute()
        except:
            erroredSheets.append(sheet)
    print('')
    print('')
    print('~~Done!~~ ')


import openpyxl
from openpyxl import *
import os
from pathlib import Path
from orderNumbers import external_order_numbers
from tqdm import tqdm

# Proceeed with the understanding that we will be running the script in the current directory of all the emails
# and the change log.  The path can always be changed.

Folder_path = Path("/home/zach/python-projects/Excel-Hyperlinks/files/")

ext_order_numbers = external_order_numbers

class emailDirectory:
    
    def __init__(self, directoryPath):
        self.directoryPath = directoryPath

    def listFiles(self):
        order_list = []
        files = os.listdir(self.directoryPath)
        # loop through every file in the directory
        for file in files:
            # find files that are .msg types and format
            if file.endswith('.msg'):
                # go through each file and match it to either the internal order number or external order number
                order_list.append(file)
            else:
                pass
        # return a list of pairs.  We want both the order number and the location for the link to work in excel
        return order_list


    def listFilePath(self):
        order_list_location = []
        files = os.listdir(self.directoryPath)
        pathToDir = '/home/zach/python-projects/Excel-Hyperlinks/files/'
        print(files)
        for file in files:
            if file.endswith('.msg'):
                order_list_location.append(os.path.join(pathToDir, file))
        return order_list_location


    def zipFilesAndPath(self):
        file_and_path = zip(self.listFiles(), self.listFilePath())
        list_file_and_path = list(file_and_path)
        return list_file_and_path 

class emailMessage:
     
    def __init__(self, file):
        self.file = file
       
    def getSplitMessage(self):
        message_base = os.path.basename(self.file)
        split_base = os.path.splitext(message_base)
        message_no_ext = split_base[0].split()
        return message_no_ext

    def matchInternalOrder(self):
        for word in self.getSplitMessage():
            if word[:3].lower() == 'ajh' and len(word) < 11:
                if len(word) == 9:
                    return word
                else:
                    new_word = word[:3] + "0" + word[3:]
                    return new_word
            else:
                pass

    def matchExternalOrder(self):
        ext_order_numbers = ['KLJH', 'AJHYN', 'OPJD']
        for word in self.getSplitMessage():
            for number in ext_order_numbers:
                if word[:5] == number:
                    return word
                elif word[:4] == number:
                    return word
                else:
                    pass


def linkFiles(workbook, direc):
    sheet = workbook.active
    for dir_value in direc:
        link_number = emailMessage(dir_value[0])
        if link_number.matchInternalOrder():
            linked_val = link_number.matchInternalOrder()
        elif link_number.matchExternalOrder():
            linked_val = link_number.matchExternalOrder()
        link_location = dir_value[1]
        currentRow = 2
        for value in sheet.iter_rows(min_row=2, values_only=True):
            if value[1] == linked_val:
                sheet.cell(row=currentRow, column=2).hyperlink = link_location
                sheet.cell(row=currentRow, column=2).style = 'Hyperlink'
                currentRow += 1
            else:
                currentRow += 1
    workbook.save(filename='/home/zach/python-projects/Excel-Hyperlinks/workbooks/Change Log Updated2.xlsx')


userDirectory = emailDirectory(Folder_path) 
directory = userDirectory.zipFilesAndPath()
print(directory)
book = load_workbook(filename='/home/zach/python-projects/Excel-Hyperlinks/workbooks/Change Log.xlsx')

linkFiles(workbook=book, direc = directory)


import openpyxl
from openpyxl import *
import os
from pathlib import Path
from orderNumbers import external_order_numbers
from tqdm import tqdm

# Proceeed with the understanding that we will be running the script in the current directory of all the emails
# and the change log.  The path can always be changed.

class emailDirectory:
    
    def __init__(self, directoryPath):
        self.directoryPath = directoryPath

    def listFiles(self):
        order_list = []
        files = os.listdir(self.directoryPath)
        # Loop through every file in the directory
        for file in files:
            # Find files that are .msg types and format
            if file.endswith('.msg'):
                order_list.append(file)
            else:
                pass
        return order_list


    def listFilePath(self):
        # Loop through the directory and append the paths to the file to a list.
        order_list_location = []
        files = os.listdir(self.directoryPath)
        pathToDir = '/home/zach/python-projects/Excel-Hyperlinks/files/'
        for file in files:
            if file.endswith('.msg'):
                order_list_location.append(os.path.join(pathToDir, file))
        return order_list_location


    def zipFilesAndPath(self):
        # Create a tuple of our order list and order list location.
        file_and_path = zip(self.listFiles(), self.listFilePath())
        list_file_and_path = list(file_and_path)
        return list_file_and_path 

class emailMessage:
     
    def __init__(self, file):
        self.file = file
       
    def getSplitMessage(self):
        # Create the base name of our file
        message_base = os.path.basename(self.file)
        # Split the base from the extension type (i.e. separate our the .msg, .pdf, etc...)
        split_base = os.path.splitext(message_base)
        message_no_ext = split_base[0].split()
        return message_no_ext

    def matchInternalOrder(self):
        # Internal order numbers are always of the type 'ajhxxxxxx'.
        for word in self.getSplitMessage():
            # We can't control how the email message will be type out.  This tries to catch the case where people will
            # type 'ajh xxxxx' or 'ajhxxxxxx'.
            if word[:3].lower() == 'ajh' and len(word) < 11:
                # If the len is equal to 9 and starts with 'ajh' then we know that it must have no spaces in it.
                if len(word) == 9:
                    return word
                else:
                    # If the length is not nine and is still less than 11 then we know that there is a space, so we need
                    # to fix that.
                    new_word = word[:3] + "0" + word[3:]
                    return new_word
            else:
                pass

    def matchExternalOrder(self):
        ext_order_numbers = ['KLJH', 'AJHYN', 'OPJD']
        # External order numbers can have two different lengths, 4 or 5, but they are generally fully type out and on
        # differently typed like they are in internal order numbers.
        for word in self.getSplitMessage():
            for number in ext_order_numbers:
                if word[:5] == number:
                    return word
                elif word[:4] == number:
                    return word
                else:
                    pass


def linkFiles(workbook, direc):
    # Activate the sheet in our workbook.
    sheet = workbook.active
    # Loop through the directory and create an instance of the emailMessage class.
    for dir_value in direc:
        # [0] in dir_value is our order number, [1] is our file location of the file.
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
    workbook.save(filename='/home/zach/python-projects/Excel-Hyperlinks/workbooks/Change Log Updated.xlsx')

# Specify the path to the source of our email files to link to our workbook.
Folder_path = Path("/home/zach/python-projects/Excel-Hyperlinks/files/")

# Speciy the external order numbers that we will match.
ext_order_numbers = external_order_numbers

# Create an instance of our emailDirectory class.
userDirectory = emailDirectory(Folder_path)

# The zipFilesAndPath method will createa a tuple of the order numbers and the path to the file.
userDirectoryFiles = userDirectory.zipFilesAndPath()

# Open up our workbook specifying the path to it via openpxl method.
book = load_workbook(filename='/home/zach/python-projects/Excel-Hyperlinks/workbooks/Change Log.xlsx')

# call function with our open workbook and our directory class.
linkFiles(workbook=book, direc = userDirectoryFiles)


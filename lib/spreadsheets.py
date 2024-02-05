# Functions related to reading and writing data from spreadsheets

import os #, re, shutil, csv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
#from pathlib import Path

'''Version 1.0'''

def getSpreadsheetValues(filename):
    ''' Gets spreadsheet by name and returns the spreadsheet as a worksheet and a list of column headings '''
    #path = os.path.join('data', filename) 
    wb = load_workbook(filename)
    
    sheet = wb.worksheets[0]
    values={}
    
    for col in sheet.columns:
        #column = [cell.value for cell in col if cell.value is not None]
        column = [cell.value if cell.value is not None else "" for cell in col]
        
        if len(column) > 0 and column.count("") != len(column):
            values[str(column[0]).strip()] = column[1:]
            
    return (values)


def getFileList(myDir):
    ''' Get a list of xlsx files in the given directory '''
    return [file for file in myDir.glob("[!~.]*.xlsx")]


def createSpreadsheetWithValues(path, filename, filenameExtra, values, newValues, filter = False, min = ''):
    ''' print out a new spreadsheet with the supplied values (columns in newValues with the same title replace the original version in values, can also use filter columns to replace values within a column rather than an entire column)'''
    
    wb = Workbook()
    newSheet = wb.active
    
    col = 1
    maxRow = 0
    
    #print(values)
    #print(newValues)
    
    for title, column in values.items():
        newSheet.cell(1, col, title).font = Font(bold=True)
    
        row = 2
        
        if title in newValues.keys():
            column = newValues[title]
            #print("Length of " + title + " column: " + str(len(column)))

        if filter:
            filteredColumn = zip(column, filter)           

            for filteredRow in filteredColumn:
                #print(str(row[1]) + ": " + str(x) + ", " + str(y))
                # To do: this needs fixing as setting a minimum value won't work
                if (min and filteredRow[1]) or not min:
                    newSheet.cell(row, col, filteredRow[0])
                    row+=1
        else:
            for cell in column:
                newSheet.cell(row, col, cell)
                row+=1  

        if row > maxRow:
            maxRow = row

        #print(title + ": " + str(row))    
        col+=1 

    #print(maxRow)
    if maxRow > 2:   
        if not os.path.exists(path):
            os.makedirs(path)

        newFilename = os.path.splitext(os.path.basename(filename))[0] + "_" + str(filenameExtra) + os.path.splitext(os.path.basename(filename))[1]
        newFile = os.path.join(path, newFilename)  
        wb.save(newFile)

        print("New file " + newFile)
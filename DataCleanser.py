import sys, os, re, requests, time, random
sys.path.append("C:\\Users\\flawrence\\Documents\\Projects\\pythonLibs\\ctd-python-libs")
import spreadsheets as s
from pathlib import Path
from datetime import date
from prompt_toolkit.shortcuts import checkboxlist_dialog, radiolist_dialog, button_dialog, input_dialog
from bs4 import BeautifulSoup
from faker import Faker


def getSettingsInputFromUser(total):
    settingsText = "Do you want to use the following settings for all " + str(total) + " spreadsheets?"

    saveSettings = button_dialog(
    title='Bulk or Individual Settings',
    text=settingsText,
    buttons=[
        ('Yes', True),
        ('No', False)]).run() 
    
    return saveSettings
    
def getSpreadsheetInputFromUser(columns, spreadsheetTitle = ''):
    patterns = {}

    if spreadsheetTitle != '':
        dialogTitle = "Column Selector: " + spreadsheetTitle
    else:
        dialogTitle = "Column Selector"
    
    columnResultsArray = checkboxlist_dialog(
        title=dialogTitle,
        text="Which columns would you like to replace the text in?",
        values = [(heading, heading) for heading in columns]).run()

    if columnResultsArray:
        columns = dict([(key, "") for key in columnResultsArray])

        for column in columnResultsArray:
            result = radiolist_dialog(
                title="Replacement Type",
                text="Replace " + column + " as:",
                values=[
                    ("fullname", "Fullname"),
                    ("first", "First names"),
                    ("initials", "Initials"),
                    ("mixed", "Mix of first names and initials"),
                    ("surname", "Surname"),
                    ("quicktext", "Text (quick)"),
                    ("wikitext", "Text (wikipedia)"),
                    ("job", "Occupation"),
                    ("addy", "Address")
                ]
            ).run() 
            columns[column] = result
    else:
        exit()
    
    if "quicktext" in columns.values() or "wikitext" in columns.values():
        for columnName, type in columns.items():
            if "text" in type:
                patterns[columnName] = getPattern(columnName)

    return (columns, patterns)

def getPattern(columnName, retry = False):
    #print("Getting pattern")

    if retry:
        titleText = 'Regex pattern for "' + columnName + '" not valid. Re-enter pattern or leave blank'
    else:
        titleText = 'Include pattern in text for "' + columnName + '"'


    pattern = input_dialog(
        title=titleText,
        text='Enter regex pattern of text to include from original text or leave blank:').run()

    if pattern != '': 
        try:            
            return re.compile(pattern)                         
        except re.error:
            return(getPattern(columnName, True))
    else:
        return pattern 

def includeFromOriginalEntry(pattern, originalText, newText):
    return([newRow[:-len("[Replacement]")] + " ".join(pattern.findall(oldRow)) + " [Replacement]" if re.search(pattern, oldRow) else newRow for oldRow, newRow in zip(originalText, newText)])

def createNewEntries(filename, sheet, columnsToRedact, patterns):
    replacementColumns = {}

    #print(patterns)

    for columnName, type in columnsToRedact.items():
        originalColumn = sheet[columnName]
        while originalColumn and originalColumn[-1] == '':
            originalColumn = originalColumn[:-1]

        #print(columnName + ": " + type)
        #print("Length of column: " + str(len(column)))

        identifier = filename + "_" + columnName

        if type == "surname":
            replacementColumns[columnName] = newNameColumnGenerator(len(originalColumn), "surname")
        elif type == "first":
            replacementColumns[columnName] = newNameColumnGenerator(len(originalColumn), "first")
        elif type == "initials":
            replacementColumns[columnName] = newNameColumnGenerator(len(originalColumn), "initials")
        elif type == "mixed":
            replacementColumns[columnName] = newNameColumnGenerator(len(originalColumn), "mixed")
        elif type == "fullname":
            replacementColumns[columnName] = newNameColumnGenerator(len(originalColumn), "full")
        elif type == "quicktext":
            replacementColumns[columnName] = newTextColumnGenerator(identifier, len(originalColumn))
        elif type == "wikitext":
            replacementColumns[columnName] = newTextColumnGenerator(identifier, len(originalColumn), True)
        elif type == "addy":
            replacementColumns[columnName] = newAddressColumnGenerator(len(originalColumn))
        elif type == "job":
            replacementColumns[columnName] = newJobColumnGenerator(len(originalColumn))
        else:
            raise ValueError("Column type " + type + " not recognised")
        
        if "text" in type and columnName in patterns.keys():
            pattern = patterns[columnName]
            if pattern != '':
                replacementColumns[columnName] = includeFromOriginalEntry(pattern, originalColumn, replacementColumns[columnName])

            #print(replacementColumns[columnName])    

    return replacementColumns

def newWikiTextEntry(sleep = False):
    if sleep:
        time.sleep(0.5)
    url = requests.get("https://en.wikipedia.org/wiki/Special:Random")
    soup = BeautifulSoup(url.content, "html.parser")
    title = soup.find(class_="firstHeading").text


    randomPage = "https://en.wikipedia.org/wiki/" + title
    #print(randomPage)
    url2 = requests.get(randomPage)
    soup2 = BeautifulSoup(url2.content, "html.parser")
    content = soup.find(id="mw-content-text")
    for p in content.find_all('p'):
        if p.text.strip() != "":
            newText = p.text.strip()
            newText = re.sub('\[\d+]', '', newText) + " Ref: Wikipedia - " + title
            return newText
    
    return "Perfer et obdura, dolor hic tibi proderit olim - Ovid, The Amores"

def newQuickTextEntry(count = 1):
    fake = Faker()
    return [fake.paragraph() + " [Replacement]" for i in range(count)]

def nameAndInitialsEntry(fake, type):
    numOfNames = random.choices([1, 2, 3], weights=(60, 30, 10))
    name = ""
    for i in range(numOfNames[0]):
        toInitial = random.randrange(2)

        if (toInitial == 1 and type != "initials") or type == "first":
            name += fake.first_name()
        else:
            name += (fake.first_name()[1]).upper()
        name += " "   
    return name.strip() 

def newNameColumnGenerator(count = 1, type = 'full'):
    fake = Faker()
    if type == 'full':
        return [fake.name() + " [Replacement]" for i in range(count)]
    elif type == 'surname':
        return [fake.last_name() + " [Replacement]" for i in range(count)]
    elif type == "first" or type == "initials" or type == "mixed":
        return [nameAndInitialsEntry(fake, type) + " [Replacement]" for i in range(count)]
    else:
        raise ValueError("Name type " + type + " not recognised")

def newAddressColumnGenerator(count = 1):
    fake = Faker('en_GB')
    return [(fake.address()).replace("\n", ", ") + " [Replacement]" for i in range(count)]

def newJobColumnGenerator(count = 1):
    fake = Faker()
    return [fake.job() + " [Replacement]" for i in range(count)]

def newTextColumnGenerator(identifier = '', count = 1, wiki = False):
    #print("Length of column: " + str(count))
    if wiki and count > 1:
        return [identifier + "_" + str(i + 1) + ": " + newWikiTextEntry(True) + " [Replacement]" for i in range(count)]
    elif wiki:
        return [newWikiTextEntry() for i in range(count)]    
    else:
        return [identifier + "_" + str(i + 1) + ": " + textEntry for i, textEntry in enumerate(newQuickTextEntry(count))]

def outputNewSheet(file, replacements, patterns):
        sheet = s.getSpreadsheetValues(file)
        newValues = createNewEntries(os.path.splitext(os.path.basename(file))[0], sheet, replacements, patterns)

        #print(newValues)
        print("Processing File " + str(index + 1))
        s.createSpreadsheetWithValues(Path("C:\\Users\\flawrence\\Documents\\Projects\\ctd-spreadsheet-cleanser\\data"), file, "cleaned", sheet, newValues)


'''
print("Surname:")
print(newNameColumnGenerator(2, "surname"))
print("First name:")
print(newNameColumnGenerator(2, "first"))
print("Initials:")
print(newNameColumnGenerator(2, "initials"))
print("Mixed:")
print(newNameColumnGenerator(2, "mixed"))
print("Full name:")
print(newNameColumnGenerator(2, "full"))
print("Text (Quick):")
print(newTextColumnGenerator("Test", 2))
print("Text (Wiki):")
print(newTextColumnGenerator("Test", 1, True))
print("Address:")
print(newAddressColumnGenerator(2))
print("Occupation:")
print(newJobColumnGenerator(2))

'''

files = s.getFileList(Path("C:\\Users\\flawrence\\Documents\\Projects\\ctd-data-redact-by-date\\data"))

if len(files) > 1:
    bulk = getSettingsInputFromUser(len(files))
else:
    bulk = True

userInputBySheet = {}
if bulk:
    spreadsheetValues = s.getSpreadsheetValues(files[0])
    userInput, patterns = getSpreadsheetInputFromUser(list(spreadsheetValues.keys()))
else:
    for index, file in enumerate(files):
        spreadsheetValues = s.getSpreadsheetValues(file)
        spreadsheetName = os.path.splitext(os.path.basename(file))[0] + " (" + str(index + 1) + "/" + str(len(files)) + ")"
        userInputBySheet[os.path.splitext(os.path.basename(file))[0]] = getSpreadsheetInputFromUser(list(spreadsheetValues.keys()), spreadsheetName)



if len(userInputBySheet) == 0:
    for index, file in enumerate(files):
        outputNewSheet(file, userInput, patterns)
else:
    for index, file in enumerate(files):
        replacement, patterns = userInputBySheet[os.path.splitext(os.path.basename(file))[0]]
        outputNewSheet(file, replacement, patterns)
    
''' '''




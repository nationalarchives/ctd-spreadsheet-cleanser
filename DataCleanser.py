import sys, os, re, requests, time, random
sys.path.append("C:\\Users\\flawrence\\Documents\\Projects\\pythonLibs\\ctd-python-libs")
import spreadsheets as s
from pathlib import Path
from datetime import date
from prompt_toolkit.shortcuts import checkboxlist_dialog, radiolist_dialog, button_dialog, input_dialog
from bs4 import BeautifulSoup
from faker import Faker


def getPathsFromUser(retry = False):
    inputFolder = input_dialog(
        title="Set Input Folder",
        text='Enter the full path to the folder holding the spreadsheets to be processed:').run()   

    if inputFolder == None:
        return (None, None)
    elif not(os.path.exists(inputFolder)):
        raise FileNotFoundError(inputFolder + " not found")

    outputFolder = input_dialog(
        title="Set Output Folder",
        text='Enter the full path to the folder where the processed spreadsheets should be saved:').run()   

    if outputFolder == None:
        return (None, None)
    elif not(os.path.exists(outputFolder)):
        raise FileNotFoundError(outputFolder + " not found")
    else:
        return(inputFolder, outputFolder)



def getSettingsInputFromUser(total):
    settingsText = "Do you want to use the following settings for all " + str(total) + " spreadsheets?"

    saveIdSettings = button_dialog(
    title='Bulk or Individual Settings',
    text=settingsText,
    buttons=[
        ('Yes', True),
        ('No', False)]).run() 
    
    return saveIdSettings

def getIdentifierInputFromUser(count, retry = False, retryMessg = ''):
    
    dialogueText = "How many digits? (use numbers only)"
    if retryMessg != '':
        dialogueText += " - " + retryMessg
    
    length = input_dialog(
    title='Identifier Format 1/2',
    text=dialogueText,
    ).run()

    if length != None:
        length = length.strip()
    else:
        return (None, None)

    if length and length.isdigit(): 
        uniqueId = button_dialog(
        title='Identifiers 2/2',
        text="Should the identifiers be unique or repeated?",
        buttons=[
            ('Unique', True),
            ('Repeated', False)]).run()
    else:
        return getIdentifierInputFromUser(True)

    identifierSettings = (int(length), uniqueId)
    
    return identifierSettings

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
                    ("addy", "Address"),
                    ("id_num", "Identifier (numerical)")
                ]
            ).run() 
            columns[column] = result
    else:
        exit()
    
    if "quicktext" in columns.values() or "wikitext" in columns.values():
        for columnName, type in columns.items():
            if "text" in type:
                patterns[columnName] = getPattern(columnName)

    if "number" in columns.values():
        pass

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
    linked = False
    firsttype = ""

    types = columnsToRedact.values()
    if "fullname" in types and "surname" in types and ("first" in types or "initials" in types or "mixed" in types):
        linked = True
        if "first" in types:
            firsttype = "first"
        elif "initials" in types:
            firsttype = "initials"
        elif "mixed" in types:
            firsttype = "mixed"


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
        elif type == "fullname" and not(linked):
            replacementColumns[columnName] = newNameColumnGenerator(len(originalColumn), "full") 
        elif type == "quicktext":
            replacementColumns[columnName] = newTextColumnGenerator(len(originalColumn), identifier)
        elif type == "wikitext":
            replacementColumns[columnName] = newTextColumnGenerator(len(originalColumn), identifier, True)
        elif type == "addy":
            replacementColumns[columnName] = newAddressColumnGenerator(len(originalColumn))
        elif type == "job":
            replacementColumns[columnName] = newJobColumnGenerator(len(originalColumn))
        elif type == "id_num":
            replacementColumns[columnName] = newIdentifierColumnGenerator(len(originalColumn), "numerical")
        elif type != "fullname":
            raise ValueError("Column type " + type + " not recognised")


        if "text" in type and columnName in patterns.keys():
            pattern = patterns[columnName]
            if pattern != '':
                replacementColumns[columnName] = includeFromOriginalEntry(pattern, originalColumn, replacementColumns[columnName])

            #print(replacementColumns[columnName])   

    if linked:
        firstpart = []
        secondpart = [] 
        fullnameColumn = ""       

        for columnName, type in columnsToRedact.items():
            if type == "surname":
                secondpart = replacementColumns[columnName]
            elif type == firsttype:
                firstpart = replacementColumns[columnName]

            if type == "fullname":
                fullnameColumn = columnName
        
        replacementColumns[fullnameColumn] = [fn[:-len("[Replacement]")] + sn for fn, sn in zip(firstpart, secondpart)]
        
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

def newTextColumnGenerator(count = 1, identifier = '', wiki = False):
    #print("Length of column: " + str(count))
    if wiki and count > 1:
        return [identifier + "_" + str(i + 1) + ": " + newWikiTextEntry(True) + " [Replacement]" for i in range(count)]
    elif wiki:
        return [newWikiTextEntry() for i in range(count)]    
    else:
        return [identifier + "_" + str(i + 1) + ": " + textEntry for i, textEntry in enumerate(newQuickTextEntry(count))]

def newIdentifierColumnGenerator(count = 1, type = "numerical"):
    # Need to check if numerical, letters, mix and how long

    length, unique = getIdentifierInputFromUser(count)
    if length != None and unique != None:

        #Will not create an id with a leading 0 to avoid it getting dropped
        id_string = ''.join([str(random.randint(1,9)) for i in range(length)])
        newIdentifier = int(id_string)

        if type == "numerical":

            if unique:
                identifier_set = set()
                while len(identifier_set) <= count:
                    identifier_set.add(newIdentifier)
                    newIdentifier = int(''.join([str(random.randint(0,9)) for i in range(length)]))

                identifiers = list(identifier_set)
            else:
                identifiers = [newIdentifier] * count
            
        else:
            raise ValueError("Name type " + type + " not recognised")

        return identifiers
    else:
        return None


def outputNewSheet(file, output, replacements, patterns):
        sheet = s.getSpreadsheetValues(file)
        newValues = createNewEntries(os.path.splitext(os.path.basename(file))[0], sheet, replacements, patterns)

        #print(newValues)
        print("Processing File " + str(index + 1))
        s.createSpreadsheetWithValues(Path(output), file, "cleaned", sheet, newValues)


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

'''# Testing
files = s.getFileList(Path("C:\\Users\\flawrence\\Documents\\Projects\\ctd-spreadsheet-cleanser\\data\\test"))


newIdentifiers = newIdentifierColumnGenerator(5, "numerical")
print(newIdentifiers)'''



 
#files = s.getFileList(Path("C:\\Users\\flawrence\\Documents\\Projects\\ctd-data-redact-by-date\\data"))
#files = s.getFileList(Path("C:\\Users\\flawrence\\Documents\\Projects\\ctd-spreadsheet-cleanser\\data\\test"))


''' '''

inputFolder, outputFolder = getPathsFromUser()

if outputFolder != None:

    files = s.getFileList(Path(inputFolder))

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
            outputNewSheet(file, outputFolder, userInput, patterns)
    else:
        for index, file in enumerate(files):
            replacement, patterns = userInputBySheet[os.path.splitext(os.path.basename(file))[0]]
            outputNewSheet(file, outputFolder, replacement, patterns)
        





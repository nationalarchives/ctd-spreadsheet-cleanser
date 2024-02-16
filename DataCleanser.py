import os, re, requests, time, random
from lib import spreadsheets as s
from pathlib import Path
#from datetime import date
from prompt_toolkit.shortcuts import checkboxlist_dialog, radiolist_dialog, button_dialog, input_dialog
from bs4 import BeautifulSoup
from faker import Faker
 
# To Do: Option to allow leading zeros on identifier numbers
# To Do: Additional identifier option of 'don't care' for unique vs repeated
# To Do: Allow option for alphabetic identifiers, alphanumeric identifiers and identifiers matching particular pattern
# To Do: Check for repetition in the column headings and deal with it
# To Do: specify which row to take headings from?
# To Do: Deal with multiple tabs in same spreadsheet



def getPathsFromUser():
    ''' Queries the user for, firstly, the path to the folder where the spreadsheets to be processed are and, 
        secondly, the path to the folder where the processed spreadsheets should be saved.  

        Returns tuple: path to input folder (string), path to output folder (string)    
    '''

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
    ''' Query user as to whether they want to apply the settings to all spreadsheets or to deal with the different spreadsheets differently
    
        Keyword arguments:
        total -- the total number of spreadsheets in the folder to be processed (Required)

        Returns boolean - True = process all the spreadsheets the same, Talse = process the spreadsheets separately
    '''

    settingsText = "Do you want to use the following settings for all " + str(total) + " spreadsheets?"

    saveIdSettings = button_dialog(
    title='Bulk or Individual Settings',
    text=settingsText,
    buttons=[
        ('Yes', True),
        ('No', False)]).run() 
    
    return saveIdSettings

def getIdentifierInputFromUser(columnName, retryMessg = ''):
    ''' For a given column which is being replaced by an identifier, queries the user for how long the identifier should be. 
        The function is recursive and calls itself to query the user again if the input from the user is not an integer.

        Keyword arguments:
        columnName -- the name of the column which contains an identifier to be replaced. This is included to 
                        differentiate as there might be multiple columns. (Required)
        retryMessg -- an additional message which is set when the recursion is called

        Returns tuple: length of identifier (integer), uniqueness of identifier (boolean)
    '''
    
    dialogueText = "How many digits? (use numbers only)"
    titleText1 = columnName + ": Identifier Length (1/2)"
    titleText2 = columnName + ": Identifier Uniqueness (2/2)"

    if retryMessg != '':
        dialogueText += " - " + retryMessg
    
    length = input_dialog(
    title=titleText1,
    text=dialogueText,
    ).run()

    if length != None:
        length = length.strip()
    else:
        return (None, None)

    if length and length.isdigit(): 
        uniqueId = button_dialog(
        title=titleText2,
        text="Should the identifiers be unique or repeated?",
        buttons=[
            ('Unique', True),
            ('Repeated', False)]).run()
    else:
        return getIdentifierInputFromUser(columnName, "Number entered for length of identifier not recognised, please enter an integer number.")

    identifierSettings = (int(length), uniqueId)
    
    return identifierSettings

def getSpreadsheetInputFromUser(columns, spreadsheetTitle = ''):
    ''' Query the user for which columns they want to replace values in. 
        If 'text' is chosen then call the function to get the pattern is one needs to be included

        Keyword arguments:
        columns -- list of column names (required)
        spreadsheetTitle -- name of spreadsheet (optional)
    '''
    patterns = {}
    idFormats = {}

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

    if "id_num" in columns.values():
        for columnName, type in columns.items():
            if "id_num" in type:
                idFormats[columnName] = getIdentifierInputFromUser(columnName)
    
    if "quicktext" in columns.values() or "wikitext" in columns.values():
        for columnName, type in columns.items():
            if "text" in type:
                patterns[columnName] = getPattern(columnName)

    return (columns, patterns, idFormats)

def getPattern(columnName, retry = False):
    ''' Query user for a regex pattern if they want to include a selection of the original content. 
        Function is recursive if a pattern is entered which is not valid.

        Keyword arguments:
        columnName -- the name of the column being replaced
        retry -- whether the function is being retried
    '''
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
    ''' Get any text matching the given pattern from the original text and add it to the replacement text

        Keyword arguments:
        pattern -- a regex pattern
        originalText -- original text which is being replaced
        newText -- new text which is replacing the original text
    '''
    return([newRow[:-len("[Replacement]")] + " ".join(pattern.findall(oldRow)) + " [Replacement]" if re.search(pattern, oldRow) else newRow for oldRow, newRow in zip(originalText, newText)])


def createNewEntries(filename, sheet, columnsToRedact, patterns, identifierFormats):
    ''' Generate the entries for the new spreadsheet

        Keyword arguments:
        filename -- name of the original spreadsheet (required)
        sheet -- values in the original spreadsheet (required)
        columnsToRedact -- list of columns to be redacted (required)
        patterns -- list of patterns for given columns (required)
        identifierFormat -- tuple with values for length and uniqueness for given spreadsheet (required)
    '''
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
            replacementColumns[columnName] = newIdentifierColumnGenerator(len(originalColumn), "numerical", identifierFormats[columnName])
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

def newIdentifierColumnGenerator(count = 1, type = "numerical", formats = None):
    # Need to check if numerical, letters, mix and how long
    if formats == None:
        (length, unique) = (8, True)
    else:
        (length, unique) = formats

    if length != None and unique != None:

        if type == "numerical":
            id_string = str(random.randint(1,9)) + ''.join([str(random.randint(0,9)) for i in range(length - 1)])
            newIdentifier = int(id_string)

            if unique:
                identifier_set = set()
                while len(identifier_set) < count:
                    identifier_set.add(newIdentifier)
                    newIdentifier = int(str(random.randint(1,9)) + ''.join([str(random.randint(0,9)) for i in range(length - 1)]))

                identifiers = list(identifier_set)
            else:
                identifiers = [newIdentifier] * count
            
        else:
            raise ValueError("Name type " + type + " not recognised")

        return identifiers
    else:
        return None


def outputNewSheet(file, output, replacements, patterns, idLengths):
        sheet = s.getSpreadsheetValues(file)

        #Need column name for text on dialogue as may be more than one column. 
        
        newValues = createNewEntries(os.path.splitext(os.path.basename(file))[0], sheet, replacements, patterns, idLengths)

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


''' '''

''' Main program

    Get paths from user
    Get list of files from the input directory
    Get bulk setting from user and treat spreadsheets accordingly
    If bulk changes:
        Get spreadsheet replacement settings 
    Otherwise: 
        Get spreadsheet replacement settings for each spreadsheet and store in userInputBySheet (spreadsheet file name as key)
    Generate new spreadsheet based on user replacement settings and output to specified folder
'''

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
        userInput, patterns, idLengths = getSpreadsheetInputFromUser(list(spreadsheetValues.keys()))
    else:
        for index, file in enumerate(files):
            spreadsheetValues = s.getSpreadsheetValues(file)
            spreadsheetName = os.path.splitext(os.path.basename(file))[0] + " (" + str(index + 1) + "/" + str(len(files)) + ")"
            userInputBySheet[os.path.splitext(os.path.basename(file))[0]] = getSpreadsheetInputFromUser(list(spreadsheetValues.keys()), spreadsheetName)



    if len(userInputBySheet) == 0:
        for index, file in enumerate(files):
            outputNewSheet(file, outputFolder, userInput, patterns, idLengths)
    else:
        for index, file in enumerate(files):
            replacement, patterns = userInputBySheet[os.path.splitext(os.path.basename(file))[0]]
            outputNewSheet(file, outputFolder, replacement, patterns, idLengths)
        





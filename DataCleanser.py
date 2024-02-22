import os, re, requests, time, random
from lib import spreadsheets as s
from pathlib import Path
#from datetime import date
from prompt_toolkit.shortcuts import checkboxlist_dialog, radiolist_dialog, button_dialog, input_dialog
from bs4 import BeautifulSoup
from faker import Faker
 
# To Do: Allow option for alphabetic identifiers and alphanumeric identifiers 
# To Do: Option to allow sequential numerical identifiers (starting from random number or starting from given number)
# To Do: Allow option for identifiers matching particular pattern
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
        total -- the total number of spreadsheets in the folder to be processed [int, required]

        Returns boolean: True = process all the spreadsheets the same, False = process the spreadsheets separately
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
                        differentiate as there might be multiple columns. [string, required]
        retryMessg -- an additional message which is set when the recursion is called [string, optional: default is '']

        Returns tuple: length of identifier (integer), fixed length for identifier (boolean), uniqueness of identifier (string: "unique|nonunique|repeated")
    '''
    
    dialogueText = "How many digits? (use numbers only)"
    titleText1 = columnName + ": Identifier Maximum Length (1/3)"
    titleText2 = columnName + ": Consistent Length/Leading Zeros (2/3)"
    titleText3 = columnName + ": Identifier Uniqueness (3/3)"

    if retryMessg != '':
        dialogueText += " - " + retryMessg
    
    length = input_dialog(
    title=titleText1,
    text=dialogueText,
    ).run()

    if length != None:
        length = length.strip()
    else:
        return (None, None, None)

    if length and length.isdigit(): 
        leadingZero = button_dialog(
        title=titleText2,
        text="Should the identifiers all be the same length or allow leading zeros?",
        buttons=[
            ('= Length', True),
            ('Leading 0s', False)]).run()
        

        if leadingZero != None:
            uniqueId = button_dialog(
            title=titleText3,
            text="Should the identifiers be unique or repeated?",
            buttons=[
                ('Unique', 'unique'),
                ('Non-Unique', 'nonunique'),
                ('Repeated', 'repeated')]).run()            

        else:
            return (length, None, None)

    else:
        return getIdentifierInputFromUser(columnName, "Number entered for length of identifier not recognised, please enter an integer number.")
    
    return (int(length), leadingZero, uniqueId)

def getSpreadsheetInputFromUser(columns, spreadsheetTitle = ''):
    ''' Query the user for which columns they want to replace values in. 
        If 'text' is chosen then call the function to get the pattern is one needs to be included

        Keyword arguments:
        columns -- list of column names [list, required]
        spreadsheetTitle -- filename of spreadsheet [string, optional: default is '']

        Returns tuple: dictionary of type of replacement by columnName, dictionary of patterns by columnName (if any), dictionary of identifier formats by columnName (if any)
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
        columnName -- the name of the column being replaced [string, required]
        retry -- whether the function is being retried [boolean, optional: default is False]

        Returns string: regex pattern
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
        pattern -- a regex pattern [string, required]
        originalText -- the original text which is being replaced [string, required]
        newText -- string with the new text which is replacing the original text [string, required]

        Returns string: replacement text with anything matching the pattern from the original text added 
    '''
    return([newRow[:-len("[Replacement]")] + " ".join(pattern.findall(oldRow)) + " [Replacement]" if re.search(pattern, oldRow) else newRow for oldRow, newRow in zip(originalText, newText)])


def createNewEntries(filename, sheet, columnsToRedact, patterns, identifierFormats):
    ''' Generate the entries for the new spreadsheet

        Keyword arguments:
        filename -- name of the original spreadsheet [string, required]
        sheet -- values in the original spreadsheet [dictionary, required]
        columnsToRedact -- list of columns to be redacted by column name [dictionary, required]
        patterns -- list of patterns for given columns by column name [dictionary, required]
        identifierFormat -- tuple with values for length and uniqueness for given spreadsheet [tuple, required]

        Returns dictionary: values for the new spreadsheet by columnNames
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
    ''' Retrieve the first paragraph of a random wikipedia page

        Keyword arguments:
        sleep - sleep between request calls [boolean, optional: default is False]

        Returns string: random text paragraph drawn from wikipedia or quote from Ovid if not able to get wikipedia paragraph
    '''

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
    ''' Generate a list of fake paragraphs

        Keyword arguments:
        count: number of fake paragraphs to generate [int, optional: default is 1]

        Returns list: fake text paragraphs
    '''
    fake = Faker()
    return [fake.paragraph() + " [Replacement]" for i in range(count)]


def nameAndInitialsEntry(fake, type):
    ''' Generate fake first name(s) and/or initial(s) with between one and three parts

        Keyword arguments:
        fake -- Faker seed [Faker object, required]
        type -- type of first name(s) to generate [string: "initials"|"first"|"mixed", required]

        Returns string: first name(s) and/or initial(s)
    '''
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
    ''' Generate fake name

        Keyword arguments:
        count -- how many names should be generated []int, optional: default is 1]
        type -- type of name to generate [string: "full"|"surname"|"initials"|"first"|"mixed", optional: default is "full"]

        Returns list: generated names of the specified type
    '''

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
    ''' Generate fake address (uk style)

        Keyword arguments:
        count -- how many names should be generated [int, optional: default is 1]

        Returns list: generated UK style addresses
    '''

    fake = Faker('en_GB')
    return [(fake.address()).replace("\n", ", ") + " [Replacement]" for i in range(count)]


def newJobColumnGenerator(count = 1):
    ''' Generate fake occupation

        Keyword arguments:
        count -- how many names should be generated [int, optional: default is 1]

        Returns list: generated occupations
    '''

    fake = Faker()
    return [fake.job() + " [Replacement]" for i in range(count)]


def newTextColumnGenerator(count = 1, identifier = '', wiki = False):
    ''' Generate fake text of the specified type

        Keyword arguments:
        count -- how many names should be generated [int, optional: default is 1]
        identifier -- identifier for the cell so it can be tracked during later processing [string, optional: default is ""]
        wiki -- whether the text should be extracted from wikipedia [boolean, optional: default is False]

        Returns list: generated texts
    '''

    if identifier != '':
        identifier = identifier + "_"

    #print("Length of column: " + str(count))
    if wiki and count > 1:
        return [identifier + str(i + 1) + ": " + newWikiTextEntry(True) + " [Replacement]" for i in range(count)]
    elif wiki:
        return [newWikiTextEntry() for i in range(count)]    
    else:
        return [identifier + str(i + 1) + ": " + textEntry for i, textEntry in enumerate(newQuickTextEntry(count))]

def newIdentifierColumnGenerator(count = 1, type = "numerical", formats = (8, True, "nonunique")):
    ''' Generate fake identifier

        Keyword arguments:
        count -- how many identifiers should be generated [int, optional: default is 1]
        type -- type of identifier  (currently only numerical identifiers implemented)  [string, optional: default is "numerical"]
        formats -- tuple with values for length of identifier and whether it is unique [tuple, optional: default is length 8, consistent length and non-unique]
    
        Returns list: generated identifiers
    '''

    (length, consistent, unique) = formats

    if length != None and consistent != None and unique != None:

        if type == "numerical":

            newIdentifier = generateNumber(length, consistent)

            if unique == 'unique':
                identifier_set = set()
                while len(identifier_set) < count:
                    identifier_set.add(newIdentifier)
                    newIdentifier = generateNumber(length, consistent)

                identifiers = list(identifier_set)
            elif unique == 'nonunique':
                identifiers = []
                while len(identifiers) < count:
                    identifiers.append(newIdentifier)
                    newIdentifier = generateNumber(length, consistent)                              
            else:
                identifiers = [newIdentifier] * count
            
        else:
            raise ValueError("Name type " + type + " not recognised")

        return identifiers
    else:
        return None


def generateNumber(length, consistent = True):
    ''' Generate a number given a length and whether leading zeros are allowed

        Keyword arguments:
        length -- (max) length of number to be generated [int, required]
        consistent -- consistent length or allowing leading zeros [boolean, optional: default is True (consistent length)]

        return int: number which matches the desired characteristics
    '''

    if consistent: 
        id_string = str(random.randint(1,9)) + ''.join([str(random.randint(0,9)) for i in range(length - 1)])
    else:
        id_string = ''.join([str(random.randint(0,9)) for i in range(length)])

    return int(id_string)

def outputNewSheet(file, output, replacements, patterns, idLengths):
    ''' Opens the original file to get the values, processes them to generate the values for a replacement spreadsheet and saves the new version

    Keyword arguments:
    file -- path to original spreadsheet to be processed [string, required]
    output -- path to folder where output spreadsheets are to be saved [string, required]
    replacements -- list of what type of replacements are required by column name [dictionary, required]
    patterns -- list of any patterns of test which should be copied across the original to the processed spreadsheet during a text replacement by column name [dictionary, required]
    idLengths -- list of identifier requirements by column name [dictionary, required]
    '''
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
        





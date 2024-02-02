# ctd-spreadsheet-cleanser
script to replace data in specified columns with generic text so that the spreadsheet can be used for testing without potentially sensitive real data being included

# Installation
This application relies on the spreadsheet.py module held in the ctd-python-libs repo. Download the file onto your local computer. The path to the spreadsheet.py folder is currently hard coded into the spreadsheet-cleanser so you need to edit the spreadsheet-cleanser code with the path on ytour local machine to whereever you have put the spreadsheet.py file.

The script also requires the following modules to be installed in the python environment that you are using:
* requests
* beautiful soup
* faker
* prompt_toolkit
* openpyxl (spreadsheets.py)

The script expects any spreadsheets to be converted to be held in the data folder

# Running the script
Note: The script will only replace the number of rows that are populated in the original spreadsheet.

When run from the commendline, the script will present the user with a series of options which can be navigated with the arrow keys and using the return to select, and tab to move between control areas:
* **Do you want the following setting to be used for all X spreadsheets**? (where x is the number of spreadsheets found in the data folder)

_If you select 'yes' then you will be presented with the following questions once, if 'no' then you will be asked the following for each spreadsheet_.

* **Which columns would you like to replace the text in**?
The list of columns is automatically generated from either the first spreadsheet in the folder if the spreadsheets are being generated in bulk or the specific spreadsheet it they are being treated individually. Use the arrow keys and Use the arrow keys/return to select the arrows column(s) that you want to be replaced and then tab to the okay/cancel

* **For each column selected you will be given the option to replace it with one of**:
  * Fullname
  * First names
  * Initials
  * Mix of initials and first names
  * Surname
  * Text (quick) - replaces the text with some randomised fake text
  * Text (wikipedia) - replaces the text with the first paragraph from a random wikipedia page
  * Occupation
  * Address

If you selected either Text (quick) or Text (wikipedia) then you will be given the option to enter a regex pattern. If a valid regex is entered then any match for the pattern in the original column in the spreadsheet then any matching text from each row will be appended on to the end of the generated text for that row.

A copy of the file will be generated with _cleaned appended to the file name. 



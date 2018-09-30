# ZIPGRADE

This repository contains three Excel macros which were written with the intention of improving the user experience when working with Zip Grade (a web/mobile app alternative to Scantron). 

## Files:
- FormatZipGrade.bas
- eMailWindows.bas .....(Windows)
- eMailMacOS.bas .......(MacOS)
- RDBMacOutlook.scpt ...(MacOS)

## What they do:

- FormatZipGrade.bas converts a "raw" .csv file (as exported from CANVAS) into a .csv file compatible with Zip Grade; this is particularly useful at the time one needs to upload the students' information into ZipGrade. 
- eMailWindows.bas takes a .csv file exported from ZipGrade (master file) and sends emails with feedback to each of the students in that file. The student's information is attached as .csv file composed by a single row of the master file. 
- eMailMacOS.bas does the same for mac, except that for this case the attached file is in .xlsx format
- RDBMacOutlook.scpt is a script necessary to send emails using Outlook for MacOS.

The two last files are based on [Ron De Bruin](http://www.rondebruin.nl) previously written code.

## Requirements:

- Microsoft Outlook installed and working on the local machine with the email account of interest.
- In ZipGrade, when entering the students' info, make sure that their emails go in the field: External ID.
- FormatZipGrade.bas requires a .csv file as exported by CANVAS
- eMailWindows.bas (or eMailMacOS.bas if you are using a mac) requires a .csv file exported using ZipGrade that contains a column called External Id with the emails of the students.

## Known limitations:

- eMailMacOS.bas works only for Excel 2016.
- eMailWindows.bas and eMailMacOS have been tested only for the standard format in Export as .csv (Zipgrade)

## How to install:

- First, download the repository (as a .zip file) and unzip it into a known location.

[Imgur Image](https://i.imgur.com/YXLWCuC.png)

- Go to Excel and record an empty macro: At the bottom, click on the icon next to the word Ready

[Imgur Image](https://i.imgur.com/7eQNMks.png)

- When prompted to save the macro, select to store it in the Personal Macro Workbook. This is important because it will allow installing our macros in that location which is hidden by default.

[Imgur Image](https://i.imgur.com/QGXl2kg.png)

- Then go again to the icon next to "Ready" (by now it has become a square) and click on it to stop recording.
- Make the Developer tab visible. (Office 2013 and up) Go to FILE/Options/Customize Ribbons (for Windows)  or Excel/Preferences/Ribbon & Toolbar (for Mac OS) and check the box Developer on the panel to the right.
- Go to the Developer Tab and click on the Visual Basic icon (upper left).  That will open the Visual Basic Editor 
- At the top left panel, right-click on VBA Project(PERSONAL.XLSB)/Modules 

[Imgur Image](https://i.imgur.com/mfjyjKu.png)

- Select import file in the contextual menu, then navigate to the location where the unzipped repository was saved and select one of the .bas files there. Repeat this step to upload the second one.

### Windows users:

- In the Visual Basic Editor, go to Tools/References 

[Imgur Image](https://i.imgur.com/RAM3eZN.png)

and check the box "Microsoft Outlook 16.0 Object Library" (the number could be different depending on the given version of Excel).

### MacOS users:

- Open a Finder Window
- Hold the Alt key when you press on Go in the Finder menu bar
- Click on Library
- Click on Application Scripts (if it exists; if not create this folder)
- Click on com.microsoft.Excel (if it exists; if not create this folder) note: Capital letter E
- Copy the script file from the download (RDBMacOutlook.scpt) inside this folder

## How to use:

With the file of interest already opened in Excel, go to VIEW/Macros and select the one needed.


## Once using it, how to avoid entering the instructor's email address repeatedly (Windows users):

- Go to the Developer tab and open the Visual Basic Editor
- In the left panel, double click on the eMail icon. The code will show up in the panel to the right.
- Scroll down and search for instructions at: 
######### MAKE YOUR EMAIL ACCOUNT PERMANENT HERE  ###########
- There you will also find an explanation on how to make changes to the subject and the body of the message.
- Save your changes before leaving.

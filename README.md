# zipgrade

This repository contains two Excel macros (FormatZipGrade.bas and eMail.bas) which were written with the intention of improving the user experience when working with Zip Grade (a web/mobile app alternative to Scantron).
DISCLOSURE: Some code has been taken from resources online. 
What they do:
- FormatZipGrade.bas converts a "raw" .csv file (as exported from CANVAS) into a .csv file compatible with Zip Grade; this is particularly useful at the time one needs to upload the students' information into ZipGrade. 
- eMail.bas takes a .csv file exported from ZipGrade and sends emails with feedback to each of the students in the file. 

REQUIREMENTS:
- Have Microsoft Outlook installed and working on the local machine.
- In ZipGrade, when entering the students' data, make sure that their emails go in the field External ID.
- (Tested only on Windows so far).

How to install:
- First, download the repository (as a .zip file) and unzip it into a known location.
- Go to Excel and record an empty macro (Green bottom ribbon, click the icon next to the word READY). When prompted to save the macro, select to store it in the Personal Macro Workbook. This is important because it will allow installing our macros in that location which is hidden by default.
- Make the Developer tab visible. (Office 2013 and up) Go to FILE/Options/Customize Ribbons and check the box Developer on the panel to the right.
- Go to the Developer Tab and click on the Visual Basic icon (upper left).  That will open the Visual Basic Editor 
- At the top left panel, right click on VBA Project(PERSONAL.XLSB)/Modules and select import file in the contextual menu, then navigate to the location where the unzipped repository was saved and select one of the .bas files there. Repeat this step to upload the second one.
- In the Visual Basic Editor, go to Tools/References and check the box "Microsoft Outlook 16.0 Object Library" (the number could be different depending on the given version of Excel).

How to use:
In Excel, go to VIEW/Macros and select the one needed.

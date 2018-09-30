Attribute VB_Name = "eMail_MacOS"

Sub eMailMac()
         'Based in modules and functions created by Ron De Bruin (http://www.rondebruin.nl).
         'This Sub uses the dafult email account setup in Outlook for mac. Does not work with
         ' Excel versions previous to Excel 2016.
    Dim rng As Range
    Dim Ash As Worksheet
    Dim Cws As Worksheet
    Dim Rcount As Long
    Dim Rnum As Long
    Dim FilterRange As Range
    Dim FieldNum As Integer
    Dim NewWB As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim FolderName As String
    Dim Folderstring As String
    Dim strbody As String

    Dim lastRow As Long, lastCol As Long
    Dim aCell As Range
    Dim I As Integer
    Dim delRows As Integer
    Dim strEmail As String
    Dim FoundAccount As Boolean
    Dim wantDialogBox As Boolean
        
  
    ' #####################    USE THIS PART TO CUSTOMIZE ANY MESSAGE       #####################################
      'Create the body text in the strbody string. The first and last line are used to set the font and font size
        strbody = "<FONT size=""3"" face=""Calibri"">"
    
        strbody = strbody & "Hi there" & "<br>" & "<br>" & _
            "Please, find attached your feedback" & "<br>" & _
            "Regards"
    
        strbody = strbody & "</FONT>"
  '##################################################################################################
  
  
  
   'Make folder in the Office folder if it not exists and create the path
       FolderName = "TempSaveFolder"
       Folderstring = CreateFolderinMacOffice2016(NameFolder:=FolderName)
     
    'Set filter sheet, you can also use Sheets("MySheet")
    Set Ash = ActiveSheet
    
    'Find the last non-blank cell in column A(1)###
    lastRow = Ash.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1 ###
    lastCol = Ash.Cells(1, Columns.Count).End(xlToLeft).Column

   'Find the column with email addresses  ###
   Set aCell = Ash.Range(Cells(1, 1), Cells(1, lastCol)).Find("External ID")
    
'Deleting rows that do not contain an email address. Bottom-up because the
'remaining cells shift up when one is deleted, creating bugs ###
   delRows = 0
   For I = lastRow To 2 Step -1
        If Not Ash.Cells(I, aCell.Column).Value Like "?*@?*.?*" Then
            Rows(I).EntireRow.Delete
            delRows = delRows + 1
        End If
   Next I
    lastRow = lastRow - delRows
    
    'Set filter range and filter column (column with e-mail addresses) ###Mod###
    Set FilterRange = Ash.Range(Cells(1, 1), Cells(lastRow, lastCol))
    FieldNum = aCell.Column    'Filter column = B because the filter range start in column A. ###Mod###
    
    'Add a worksheet for the unique list and copy the unique list in A1
    Set Cws = Worksheets.Add
    FilterRange.Columns(FieldNum).AdvancedFilter _
            Action:=xlFilterCopy, _
            CopyToRange:=Cws.Range("A1"), _
            CriteriaRange:="", Unique:=True

    'Count of the unique values + the header cell
    Rcount = Application.WorksheetFunction.CountA(Cws.Columns(1))

    With Application
            .EnableEvents = False
            .ScreenUpdating = False
    End With


   'If there are unique values start the loop
    If Rcount >= 2 Then
        For Rnum = 2 To Rcount

            'If the unique value is a mail addres create a mail
            If Cws.Cells(Rnum, 1).Value Like "?*@?*.?*" Then

                'Filter the FilterRange on the FieldNum column
                FilterRange.AutoFilter Field:=FieldNum, _
                                       Criteria1:=Cws.Cells(Rnum, 1).Value

                'Copy the visible data in a new workbook
                With Ash.AutoFilter.Range
                    On Error Resume Next
                    Set rng = .SpecialCells(xlCellTypeVisible)
                    On Error GoTo 0
                End With

                Set NewWB = Workbooks.Add(xlWBATWorksheet)

                rng.Copy
                With NewWB.Sheets(1)
                    .Cells(1).PasteSpecial Paste:=8
                    .Cells(1).PasteSpecial Paste:=xlPasteValues
                    .Cells(1).PasteSpecial Paste:=xlPasteFormats
                    .Cells(1).Select
                    Application.CutCopyMode = False
                End With

                'Create a file name (& Application.PathSeparator & FolderName)
                TempFilePath = Folderstring & Application.PathSeparator
                TempFileName = "Feedback_" & Ash.Parent.Name _
                             & " " & Format(Now, "dd-mmm-yy h-mm-ss")
                
                FileExtStr = ".xlsx": FileFormatNum = 51
                
                'Save, Mail, Close and Delete the file ()
                
                With NewWB
                    .SaveAs TempFilePath & TempFileName _
                          & FileExtStr, FileFormat:=FileFormatNum
                    On Error Resume Next
                    
                    
           'Call the MailWithMacOutlook2016Workbook function to create the mail
            'When you use more mail addresses separate them with a ,
            'Change yes to no in the displaymail argument to send directly
            'Look in Outlook>Preferences for the type and name of the account that you want to use
            'If accounttype is empty it will use the default mail account, accounttype can be pop or imap
            'Note: It will use the signature of the account that you choose
            MailWithMacOutlook2016Workbook _
            subject:="Feedback", _
            mailbody:=strbody, _
            toaddress:=Cws.Cells(Rnum, 1).Value, _
            ccaddress:="", _
            bccaddress:="", _
            displaymail:="no", _
            accounttype:="", _
            accountname:="", _
            attachment:=TempFilePath & TempFileName & FileExtStr
         
            'Delete the file we just mailed
            KillFileOnMac TempFilePath & TempFileName & FileExtStr
                
On Error GoTo 0
                    .Close SaveChanges:=False
                End With

            End If

            'Close AutoFilter
            Ash.AutoFilterMode = False

        Next Rnum
        
       Application.DisplayAlerts = False
       Cws.Delete
       Application.DisplayAlerts = True
    End If

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

End Sub

Function KillFileOnMac(Filestr As String)
    Dim ScriptToKillFile As String
    Dim Fstr As String
    'Ron de Bruin, 25-June-2015
    'Delete files from a Mac. Uses AppleScript to avoid
    'the problem with long file names in Office 2011
    If Val(Application.Version) < 15 Then
        ScriptToKillFile = "tell application " & Chr(34) & _
            "Finder" & Chr(34) & Chr(13)
        ScriptToKillFile = ScriptToKillFile & _
            "do shell script ""rm "" & quoted form of posix path of " & _
            Chr(34) & Filestr & Chr(34) & Chr(13)
        ScriptToKillFile = ScriptToKillFile & "end tell"

        On Error Resume Next
        MacScript (ScriptToKillFile)
        On Error GoTo 0
    Else
        Fstr = MacScript("return POSIX path of (" & _
            Chr(34) & Filestr & Chr(34) & ")")
        On Error Resume Next
        Kill Fstr
    End If
End Function


Function CreateFolderinMacOffice2016(NameFolder As String) As String
    'Function to create folder if it not exists in the Microsoft Office Folder
    'Ron de Bruin : 8-Jan-2016
    Dim OfficeFolder As String
    Dim PathToFolder As String
    Dim TestStr As String

    OfficeFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    OfficeFolder = Replace(OfficeFolder, "/Desktop", "") & _
       "Library/Group Containers/UBF8T346G9.Office/"

    PathToFolder = OfficeFolder & NameFolder

    On Error Resume Next
    TestStr = Dir(PathToFolder, vbDirectory)
    On Error GoTo 0
    If TestStr = vbNullString Then
        MkDir PathToFolder
        'You can use this msgbox line for testing if you want
        'MsgBox "You find the new folder in this location :" & PathToFolder
    End If
    CreateFolderinMacOffice2016 = PathToFolder
End Function

Function CheckAppleScriptTaskExcelScriptFile(ScriptFileName As String) As Boolean
    'Function to Check if the AppleScriptTask script file exists
    'Ron de Bruin : 6-March-2016
    Dim AppleScriptTaskFolder As String
    Dim TestStr As String

    AppleScriptTaskFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    AppleScriptTaskFolder = Replace(AppleScriptTaskFolder, "/Desktop", "") & _
        "Library/Application Scripts/com.microsoft.Excel/"

    On Error Resume Next
    TestStr = Dir(AppleScriptTaskFolder & ScriptFileName, vbDirectory)
    On Error GoTo 0
    If TestStr = vbNullString Then
        CheckAppleScriptTaskExcelScriptFile = False
    Else
        CheckAppleScriptTaskExcelScriptFile = True
    End If
End Function

Function MailWithMacOutlook2016Workbook(subject As String, mailbody As String, _
    toaddress As String, ccaddress As String, _
    bccaddress As String, displaymail As String, _
    accounttype As String, accountname As String, _
    attachment As String)
    'Function to create a mail with the activeworkbook
    ' Ron de Bruin : 10-March-2018
    ' Change the displaymail argument to string instead of boolean
    Dim fileattachment As String
    Dim ScriptStr As String, RunMyScript As String

    'Build the AppleScriptTask parameter string
    fileattachment = attachment
    ScriptStr = subject & ";" & mailbody & ";" & toaddress & ";" & ccaddress & ";" & _
                bccaddress & ";" & displaymail & ";" & accounttype & ";" & _
                accountname & ";" & fileattachment

    'Call the RDBMacOutlook.scpt script file with the AppleScriptTask function
    RunMyScript = AppleScriptTask("RDBMacOutlook.scpt", "CreateMailInOutlook", CStr(ScriptStr))

End Function



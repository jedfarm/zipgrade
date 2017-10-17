Attribute VB_Name = "eMail"

Sub Send_Row_Or_Rows_Attachment_2()
'Working in 2000-2016
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    Dim OutApp As Object
    Dim OutMail As Outlook.MailItem  ' was Object ###
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
'Modified or added later###
    Dim lastRow As Long, lastCol As Long
    Dim aCell As Range
    Dim I As Integer
    Dim delRows As Integer
    Dim strEmail As String
    Dim FoundAccount As Boolean
    Dim wantDialogBox As Boolean
    
    Dim myAccounts As Outlook.Accounts
    Dim myAccount As Outlook.Account
    Dim tempAccount As Outlook.Account
    
    FoundAccount = False
    
    
    
    
    
    
   ' ################################ MAKE YOUR EMAIL ACCOUNT PERMANENT HERE  ################################################
    
    'If you do not want a dialog box asking all the time, set wantDialogBox to FALSE
    'then change strEmail to your email address. ###
    
    wantDialogBox = True
    strEmail = "robinson_crusoe@homealone.com"
    
    
   '##########################################################################################################################
    
    
    
    
    
    
     'Asking for a proper email address ###
     If wantDialogBox Then
        strEmail = InputBox(Prompt:="Enter a valid email address", _
        Title:="ENTER YOUR EMAIL", Default:="robinson_crusoe@homealone.com")
    
'            If strEmail = "robinson_crusoe@homealone.com" Or strEmail = vbNullString Then
'               emailFlag = False
'            Else
'               emailFlag = True
'            End If
     End If
     
    On Error GoTo cleanup
    Set OutApp = CreateObject("Outlook.Application")

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

'Find an Outlook account based on from email address  ####
    Set myAccounts = OutApp.Application.Session.Accounts
    For Each tempAccount In myAccounts
        If (tempAccount.SmtpAddress = strEmail) Then
            Set myAccount = tempAccount
            FoundAccount = True
            Exit For
        End If
    Next
If FoundAccount Then
 'Set filter sheet, you can also use Sheets("MySheet")
    Set Ash = ActiveSheet
    
    'Find the last non-blank cell in column A(1)###
    lastRow = Ash.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1 ###
    lastCol = Ash.Cells(1, Columns.Count).End(xlToLeft).Column

   'Find the column with email addresses  ###
   Set aCell = Ash.Range(Cells(1, 1), Cells(1, lastCol)).Find("External ID")
    
'Deleting rows that do not contain an email address. Bottom-up because the remaining cells shift up when one is deleted, creating bugs ###
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

                'Create a file name
                TempFilePath = Environ$("temp") & "\"
                TempFileName = "Your data of " & Ash.Parent.Name _
                             & " " & Format(Now, "dd-mmm-yy h-mm-ss")

                If Val(Application.Version) < 12 Then
                    'You use Excel 97-2003
                    FileExtStr = ".xls": FileFormatNum = -4143
                Else
                    'You use Excel 2007-2016
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
                
                'Save, Mail, Close and Delete the file
                

                With NewWB
                    .SaveAs TempFilePath & TempFileName _
                          & FileExtStr, FileFormat:=FileFormatNum
                    On Error Resume Next
                    
                    
                Set OutMail = OutApp.CreateItem(0)
                With OutMail
                    .SendUsingAccount = myAccount
                    .To = Cws.Cells(Rnum, 1).Value
                    .Subject = "Feedback"
                    .Attachments.Add NewWB.FullName
                    .Body = "Hi there. This is your feedback"
                    .Send  'Use Send for using the app or Display for debugging
                End With
                
On Error GoTo 0
                    .Close savechanges:=False
                End With

                Set OutMail = Nothing
                Kill TempFilePath & TempFileName & FileExtStr
            End If

            'Close AutoFilter
            Ash.AutoFilterMode = False

        Next Rnum
    End If



cleanup:
    Set OutApp = Nothing
    Application.DisplayAlerts = False
    Cws.Delete
    Application.DisplayAlerts = True

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
 
 Else
    MsgBox ("Email account not found")
End If
    
End Sub
                    
                    

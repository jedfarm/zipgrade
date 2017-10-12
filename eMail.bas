Attribute VB_Name = "eMail"
Sub Send_Row_Or_Rows_Attachment_1()
'Working in 2000-2016
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    Dim outApp As Object
    Dim OutMail As Object
    Dim rng As Range
    Dim Ash As Worksheet
    Dim Cws As Worksheet
    Dim Rcount As Long
    Dim Rnum As Long
    Dim FilterRange As Range
    Dim FieldNum As Integer
    Dim mailAddress As String
    Dim NewWB As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileExtStr As String
    Dim FileFormatNum As Long

    On Error GoTo cleanup
    Set outApp = CreateObject("Outlook.Application")

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    'Set filter sheet, you can also use Sheets("MySheet")
    Set Ash = ActiveSheet

    'Set filter range and filter column (column with names)
    Set FilterRange = Ash.Range("A1:H" & Ash.Rows.Count)
    FieldNum = 1    'Filter column = A because the filter range start in column A

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

            'Look for the mail address in the MailInfo worksheet
            mailAddress = ""
            On Error Resume Next
            mailAddress = Application.WorksheetFunction. _
                VLookup(Cws.Cells(Rnum, 1).Value, _
                          Worksheets("Mailinfo").Range("A1:B" & _
                                Worksheets("Mailinfo").Rows.Count), 2, False)
            On Error GoTo 0

            If mailAddress <> "" Then

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
                Set OutMail = outApp.CreateItem(0)

                With NewWB
                    .SaveAs TempFilePath & TempFileName _
                          & FileExtStr, FileFormat:=FileFormatNum
                    On Error Resume Next
                    With OutMail
                        .To = mailAddress
                        .Subject = "Test mail"
                        .Attachments.Add NewWB.FullName
                        .Body = "Hi there"
                        .Display  'Or use Send
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
    Set outApp = Nothing
    Application.DisplayAlerts = False
    Cws.Delete
    Application.DisplayAlerts = True

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub


Sub Send_Row_Or_Rows_Attachment_2()
'Working in 2000-2016
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    Dim outApp As Object
    Dim OutMail As Object
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

    On Error GoTo cleanup
    Set outApp = CreateObject("Outlook.Application")

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    'Set filter sheet, you can also use Sheets("MySheet")
    Set Ash = ActiveSheet

    'Set filter range and filter column (column with e-mail addresses)
    Set FilterRange = Ash.Range("A1:R" & Ash.Rows.Count)
    FieldNum = 2    'Filter column = B because the filter range start in column A

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
                Set OutMail = outApp.CreateItem(0)

                With NewWB
                    .SaveAs TempFilePath & TempFileName _
                          & FileExtStr, FileFormat:=FileFormatNum
                    On Error Resume Next
                    With OutMail
                        '.SendUsingAccount = "jfandino3@hccfl.edu"
                        .To = Cws.Cells(Rnum, 1).Value
                        .Subject = "Test mail"
                        .Attachments.Add NewWB.FullName
                        .Body = "Hi there"
                        .Display  'Or use Send
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
    Set outApp = Nothing
    Application.DisplayAlerts = False
    Cws.Delete
    Application.DisplayAlerts = True

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub


Public Sub New_Mail()
Dim outApp As Outlook.Application
Set outApp = New Outlook.Application

Dim oAccount As Outlook.Account
Dim oMail As Outlook.MailItem
For Each oAccount In outApp.Session.Accounts
   If oAccount = "jfand72@hccfl.edu" Then
      Set oMail = outApp.CreateItem(olMailItem)
      oMail.SendUsingAccount = oAccount
      oMail.Display
   End If
Next
End Sub

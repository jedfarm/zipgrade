Attribute VB_Name = "FormatZipGrade"
Sub FormatToZipgrade()

    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim here As Boolean
    Dim colnames
    Dim tmpArray() As String
    
    'Setting up the columns we want to keep
    colnames = Array("Student", "Section", "SIS User ID", "SIS Login ID")
    
    'Deleting the second row because it does not contain info about students
    If Trim(Cells(2, 1).Value) = "Points Possible" Then
        Rows(2).EntireRow.Delete
    End If
        
        
    'Find the last non-blank cell in column A(1)
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        
    'Deleting columns that are not of interest
     For i = lastCol To 1 Step -1
        here = False
        For j = LBound(colnames) To UBound(colnames)
            If Cells(1, i).Value = colnames(j) Then
                here = True
                Exit For
            End If
        Next j
        If Not here Then
            Columns(i).EntireColumn.Delete
        End If
    Next i
    ' Deleting any Test Student account
    If Cells(lastRow, 1).Value = "Test Student" Then
        Rows(lastRow).EntireRow.Delete
    End If
    
    'Inserting a column to the left of column B
    If Cells(1, 2).Value <> "Last Name" Then
        Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove 'or xlFormatFromRightOrBelow
    
        'Splitting the names column into two
        For i = 2 To lastRow
            If InStr(1, Range("A" & i).Value, " ") Then
                tmpArray = Split(Range("A" & i).Value, " ", 2)
                Range("A" & i).Value = tmpArray(0)
                Range("B" & i).Value = tmpArray(1)
            End If
        Next i
        'Renaming the first two columns
        Cells(1, 1).Value = "Name"
        Cells(1, 2).Value = "Last Name"
    End If
      
    
End Sub


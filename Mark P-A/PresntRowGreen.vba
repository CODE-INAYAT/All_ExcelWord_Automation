' This code make the rows green where ever "P" is there. 
Sub HighlightPresentRows()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim colNum As Long
    Dim i As Long

    ' Set your worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change if needed

    ' Find the "Attendance" column number
    On Error Resume Next
    colNum = Application.WorksheetFunction.Match("Attendance", ws.Rows(1), 0)
    On Error GoTo 0

    If colNum = 0 Then
        MsgBox "Attendance column not found in Row 1.", vbExclamation
        Exit Sub
    End If

    ' Find the last used row and column
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Loop through each row in Attendance column
    For i = 2 To lastRow
        If Trim(ws.Cells(i, colNum).Value) = "P" Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Interior.Color = RGB(198, 239, 206) ' Light green
        End If
    Next i

    MsgBox "Rows with 'P' in Attendance column are now highlighted in light green.", vbInformation
End Sub

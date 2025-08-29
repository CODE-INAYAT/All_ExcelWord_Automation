' This code marks absent when empty cell finds (Means where P is not)
Sub MarkAbsentInAttendance()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim colNum As Long
    Dim i As Long
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Find the column number for "Attendance"
    On Error Resume Next
    colNum = Application.WorksheetFunction.Match("Attendance", ws.Rows(1), 0)
    On Error GoTo 0
    
    If colNum = 0 Then
        MsgBox "Column 'Attendance' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Determine last used row in the entire sheet
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    ' Loop through all rows from 2 to lastRow
    For i = 2 To lastRow
        If IsEmpty(ws.Cells(i, colNum).Value) Then
            ws.Cells(i, colNum).Value = "A"
        End If
    Next i

    MsgBox "'A' has been marked for all empty cells in the 'Attendance' column.", vbInformation
End Sub
Sub TrimAllCellSpaces()

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim cell As Range
    Dim cellsChanged As Long

    ' 1. Validate and set the worksheet object to the active sheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This macro must be run on a standard worksheet.", vbExclamation, "Invalid Sheet Type"
        Exit Sub
    End If
    Set ws = ActiveSheet
    
    ' Turn off screen updating for a huge performance gain
    Application.ScreenUpdating = False

    ' 2. Automatically find the last row and column to define the data range
    On Error Resume Next
    lastRow = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0
    
    ' If the sheet is empty, exit gracefully
    If lastRow = 0 Or lastCol = 0 Then
        MsgBox "No data found on the active sheet to process.", vbInformation, "Sheet is Empty"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Define the range to process as the entire used area of the sheet
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' 3. Loop through each cell in the automatically determined range
    cellsChanged = 0
    For Each cell In targetRange
        ' Check if the cell contains a string (text)
        If VarType(cell.Value) = vbString Then
            ' Check if trimming is actually needed before making a change
            If cell.Value <> Trim(cell.Value) Then
                cell.Value = Trim(cell.Value)
                cellsChanged = cellsChanged + 1
            End If
        End If
    Next cell

    ' Turn screen updating back on
    Application.ScreenUpdating = True

    ' 4. Display a more informative success message
    MsgBox "Operation Complete." & vbCrLf & vbCrLf & cellsChanged & " cell(s) were trimmed.", vbInformation, "Trim Spaces"
End Sub
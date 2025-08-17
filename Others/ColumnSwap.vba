' Function for swap column position
Sub SwapColumns()
    Dim col1 As String
    Dim col2 As String
    Dim ws As Worksheet
    Dim tempRange As Range
    
    col1 = InputBox("Enter the first column letter to swap (e.g., A):", "First Column")
    If col1 = "" Then Exit Sub
    
    col2 = InputBox("Enter the second column letter to swap (e.g., B):", "Second Column")
    If col2 = "" Then Exit Sub
    
    Set ws = ActiveSheet
    
    On Error Resume Next
    Set col1Range = ws.Columns(col1)
    Set col2Range = ws.Columns(col2)
    
    If col1Range Is Nothing Or col2Range Is Nothing Then
        MsgBox "Invalid column letters entered. Please try again.", vbExclamation
        Exit Sub
    End If
    
    Set tempRange = ws.Columns(ws.Columns.Count)
    
    On Error GoTo 0

    tempRange.Value = col1Range.Value
    col1Range.Value = col2Range.Value
    col2Range.Value = tempRange.Value
    tempRange.Clear

    MsgBox "Columns " & col1 & " and " & col2 & " have been swapped.", vbInformation
End Sub
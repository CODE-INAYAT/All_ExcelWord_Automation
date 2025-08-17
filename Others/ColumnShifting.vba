' Function for shifting column position
Sub ShiftColumnPosition()
    Dim sourceCol As String
    Dim targetCol As String
    Dim ws As Worksheet
    Dim colToMove As Range
    Dim colTarget As Range

    sourceCol = InputBox("Enter the column letter to move (e.g., A):", "Source Column")
    If sourceCol = "" Then Exit Sub
    
    targetCol = InputBox("Enter the column letter to move to (e.g., B):", "Target Column")
    If targetCol = "" Then Exit Sub
    
    Set ws = ActiveSheet
    
    On Error Resume Next
    Set colToMove = ws.Columns(sourceCol)
    Set colTarget = ws.Columns(targetCol)
    
    If colToMove Is Nothing Or colTarget Is Nothing Then
        MsgBox "Invalid column letters entered. Please try again.", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo 0

    If colToMove.Column < colTarget.Column Then
        colToMove.Cut
        colTarget.Insert Shift:=xlToRight
    Else
        colToMove.Cut
        colTarget.Insert Shift:=xlToLeft
    End If
    
    MsgBox "Column " & sourceCol & " has been moved to the position of column " & targetCol & ".", vbInformation
End Sub
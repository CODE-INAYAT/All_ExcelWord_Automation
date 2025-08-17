' Function For Uppercase
Sub ConvertToUppercase()
    Dim cell As Range
    On Error Resume Next
    For Each cell In Selection
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            cell.Value = UCase(cell.Value)
        End If
    Next cell
    On Error GoTo 0

    MsgBox "All selected cells are now UPPERCASE"
End Sub

Sub UppercaseTandPUIDColumn()
    ' This macro will run on the currently ACTIVE sheet.
    Const TARGET_COLUMN_HEADER As String = "T&P UID"

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetCol As Range
    Dim dataRange As Range
    Dim cell As Range

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This macro must be run on a standard worksheet.", vbExclamation, "Invalid Sheet Type"
        Exit Sub
    End If
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False

    On Error Resume Next
    Set targetCol = ws.Rows(1).Find(What:=TARGET_COLUMN_HEADER, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If targetCol Is Nothing Then
        MsgBox "The column '" & TARGET_COLUMN_HEADER & "' could not be found on the active sheet." & vbCrLf & _
               "Please check the header name.", vbCritical, "Column Not Found"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, targetCol.Column).End(xlUp).Row

    If lastRow <= 1 Then
        MsgBox "No data found in the '" & TARGET_COLUMN_HEADER & "' column.", vbInformation, "No Data"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Set dataRange = ws.Range(ws.Cells(2, targetCol.Column), ws.Cells(lastRow, targetCol.Column))

    For Each cell In dataRange
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            cell.Value = UCase(cell.Value)
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "All text in the '" & TARGET_COLUMN_HEADER & "' column has been converted to UPPERCASE.", vbInformation, "Task Complete"
End Sub
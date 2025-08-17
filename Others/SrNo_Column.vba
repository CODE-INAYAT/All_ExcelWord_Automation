Sub AddSerialNumberColumn()

    ' --- SETTINGS ---
    Const SR_NO_HEADER As String = "Sr. No."
    ' ----------------
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim userChoice As Integer
    Dim groupColName As String
    Dim groupColNum As Long
    Dim i As Long
    Dim serialCounter As Long
    Dim currentGroup As String
    Dim previousGroup As String
    Dim serialNumbers() As Variant
    
    ' 1. Validate active sheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This macro must be run on a worksheet.", vbExclamation, "Invalid Sheet"
        Exit Sub
    End If
    Set ws = ActiveSheet
    
    ' 2. Check if data exists
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "No data found in the worksheet.", vbInformation, "No Data"
        Exit Sub
    End If
    
    ' 3. Ask user for serial number type
    userChoice = MsgBox("Do you want Special Serial Numbers?" & vbCrLf & vbCrLf & _
                        "Yes = Special (resets on group change)" & vbCrLf & _
                        "No = Normal (continuous 1,2,3...)", _
                        vbYesNoCancel + vbQuestion, "Serial Number Type")
    
    ' Handle cancel
    If userChoice = vbCancel Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 4. Prepare serial numbers array
    ReDim serialNumbers(1 To lastRow - 1, 1 To 1)
    
    ' 5. Generate serial numbers based on user choice
    If userChoice = vbNo Then
        ' Normal serial numbers (1, 2, 3, ...)
        For i = 1 To lastRow - 1
            serialNumbers(i, 1) = i
        Next i
        
    ElseIf userChoice = vbYes Then
        ' Special serial numbers (reset on group change)
        
        ' Ask for column name to group by
        groupColName = InputBox("Enter the column name to group by:" & vbCrLf & vbCrLf & _
                                "Serial numbers will restart from 1 for each new value in this column.", _
                                "Group Column Name")
        
        ' Check if user cancelled or entered blank
        If Trim(groupColName) = "" Then
            MsgBox "Operation cancelled.", vbInformation
            GoTo CleanUp
        End If
        
        ' Find the group column
        On Error Resume Next
        groupColNum = ws.Rows(1).Find(What:=groupColName, LookAt:=xlWhole, MatchCase:=False).Column
        On Error GoTo 0
        
        If groupColNum = 0 Then
            MsgBox "Column '" & groupColName & "' not found.", vbExclamation, "Column Not Found"
            GoTo CleanUp
        End If
        
        ' Generate special serial numbers
        serialCounter = 1
        previousGroup = ""
        
        For i = 2 To lastRow
            currentGroup = CStr(ws.Cells(i, groupColNum).Value)
            
            ' Check if group changed
            If currentGroup <> previousGroup Then
                serialCounter = 1
                previousGroup = currentGroup
            End If
            
            serialNumbers(i - 1, 1) = serialCounter
            serialCounter = serialCounter + 1
        Next i
    End If
    
    ' 6. Insert new column at the beginning
    ws.Columns(1).Insert Shift:=xlToRight
    
    ' 7. Add header
    ws.Cells(1, 1).Value = SR_NO_HEADER
    
    ' 8. Add serial numbers
    ws.Cells(2, 1).Resize(UBound(serialNumbers, 1), 1).Value = serialNumbers
    
    ' 9. Autofit the serial number column
    ws.Columns(1).AutoFit
    
CleanUp:
    Application.ScreenUpdating = True
    
    If userChoice = vbNo Then
        MsgBox "Normal serial numbers have been added successfully!", vbInformation, "Complete"
    ElseIf userChoice = vbYes And groupColNum > 0 Then
        MsgBox "Special serial numbers have been added successfully!" & vbCrLf & _
               "Grouped by: " & groupColName, vbInformation, "Complete"
    End If

End Sub
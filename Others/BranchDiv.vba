' Generate BranchDiv column from the Branch & Divison column

Sub CreateBranchDiv_BeforeBranch_Corrected()

    ' --- SETTINGS ---
    ' This macro will run on the currently ACTIVE sheet.
    Const BRANCH_COL_HEADER As String = "Branch"
    Const DIVISION_COL_HEADER As String = "Division"
    Const NEW_COL_HEADER As String = "BranchDiv"
    ' ----------------

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim branchColNum As Long, divColNum As Long, insertionPoint As Long
    Dim i As Long
    Dim branchVal As String, divVal As String
    
    ' This array will hold the prepared data BEFORE we modify the sheet
    Dim branchDivData() As Variant

    ' 1. Validate and set the worksheet object
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This macro must be run on a standard worksheet, not a chart sheet.", vbExclamation, "Invalid Sheet Type"
        Exit Sub
    End If
    Set ws = ActiveSheet
    
    ' Turn off screen updating for performance
    Application.ScreenUpdating = False
    
    ' 2. Find required columns and last row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then
        MsgBox "No data found on the active sheet.", vbInformation, "No Data"
        GoTo CleanUp
    End If

    On Error Resume Next
    branchColNum = ws.Rows(1).Find(What:=BRANCH_COL_HEADER, LookAt:=xlWhole, MatchCase:=False).Column
    divColNum = ws.Rows(1).Find(What:=DIVISION_COL_HEADER, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0

    If branchColNum = 0 Or divColNum = 0 Then
        MsgBox "Could not find the required columns: '" & BRANCH_COL_HEADER & "' and/or '" & DIVISION_COL_HEADER & "'.", vbExclamation, "Columns Not Found"
        GoTo CleanUp
    End If

    ' 3. *** PREPARE DATA IN MEMORY FIRST (THE CORRECT METHOD) ***
    '    Resize the array to hold the data for all rows (from 2 to lastRow)
    ReDim branchDivData(1 To lastRow - 1, 1 To 1)
    
    '    Loop through the source data and generate the new values
    For i = 2 To lastRow
        branchVal = CStr(ws.Cells(i, branchColNum).Value)
        divVal = CStr(ws.Cells(i, divColNum).Value)

        ' Apply the combination logic
        If UCase(Trim(divVal)) = "NA" Then
            branchDivData(i - 1, 1) = branchVal
        Else
            branchDivData(i - 1, 1) = branchVal & "-" & divVal
        End If
    Next i

    ' 4. *** NOW, MODIFY THE WORKSHEET ***
    '    Determine the insertion point using the original Branch column number
    insertionPoint = branchColNum
    
    '    Insert a new, blank column at that position
    ws.Columns(insertionPoint).Insert Shift:=xlToRight
    
    '    Add the header to the newly created column
    ws.Cells(1, insertionPoint).Value = NEW_COL_HEADER
    
    '    Paste the entire prepared data array into the column in one go (very fast)
    ws.Cells(2, insertionPoint).Resize(UBound(branchDivData, 1), 1).Value = branchDivData

    ' 5. Autofit the new column for readability
    ws.Columns(insertionPoint).AutoFit

CleanUp:
    ' 6. Restore screen updating
    Application.ScreenUpdating = True
    
    ' Display a success message
    If branchColNum > 0 And divColNum > 0 Then
        MsgBox "'" & NEW_COL_HEADER & "' column has been successfully created before the '" & BRANCH_COL_HEADER & "' column.", vbInformation, "Task Complete"
    End If

End Sub
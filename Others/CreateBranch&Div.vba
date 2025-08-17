Sub SplitBranchDiv_AndShiftRight()

    ' --- SETTINGS ---
    Const BRANCHDIV_COL_HEADER As String = "BranchDiv"
    Const BRANCH_COL_HEADER As String = "Branch"
    Const DIVISION_COL_HEADER As String = "Division"
    ' ----------------

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim branchDivColNum As Long
    Dim i As Long
    Dim branchDivVal As String
    Dim splitParts() As String
    Dim branchData() As Variant, divisionData() As Variant

    ' 1. Validate active sheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This macro must be run on a worksheet.", vbExclamation, "Invalid Sheet"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Application.ScreenUpdating = False

    ' 2. Find "BranchDiv" column and last row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "No data to process.", vbInformation, "No Data"
        GoTo CleanUp
    End If

    On Error Resume Next
    branchDivColNum = ws.Rows(1).Find(What:=BRANCHDIV_COL_HEADER, LookAt:=xlWhole, MatchCase:=False).Column
    On Error GoTo 0

    If branchDivColNum = 0 Then
        MsgBox "'" & BRANCHDIV_COL_HEADER & "' column not found.", vbExclamation, "Column Not Found"
        GoTo CleanUp
    End If

    ' 3. Prepare arrays for new data
    ReDim branchData(1 To lastRow - 1, 1 To 1)
    ReDim divisionData(1 To lastRow - 1, 1 To 1)

    For i = 2 To lastRow
        branchDivVal = Trim(CStr(ws.Cells(i, branchDivColNum).Value))
        If InStr(branchDivVal, "-") > 0 Then
            splitParts = Split(branchDivVal, "-")
            branchData(i - 1, 1) = splitParts(0)
            divisionData(i - 1, 1) = splitParts(1)
        Else
            branchData(i - 1, 1) = branchDivVal
            divisionData(i - 1, 1) = "NA"
        End If
    Next i

    ' 4. Insert two new columns after BranchDiv column (this shifts existing columns to the right)
    ws.Columns(branchDivColNum + 1).Insert Shift:=xlToRight
    ws.Columns(branchDivColNum + 2).Insert Shift:=xlToRight

    ' 5. Add headers
    ws.Cells(1, branchDivColNum + 1).Value = BRANCH_COL_HEADER
    ws.Cells(1, branchDivColNum + 2).Value = DIVISION_COL_HEADER

    ' 6. Paste data
    ws.Cells(2, branchDivColNum + 1).Resize(UBound(branchData, 1), 1).Value = branchData
    ws.Cells(2, branchDivColNum + 2).Resize(UBound(divisionData, 1), 1).Value = divisionData

    ' 7. Autofit new columns
    ws.Columns(branchDivColNum + 1).AutoFit
    ws.Columns(branchDivColNum + 2).AutoFit

CleanUp:
    Application.ScreenUpdating = True

    If branchDivColNum > 0 Then
        MsgBox "'Branch' and 'Division' columns have been inserted after 'BranchDiv', and existing columns were shifted right.", vbInformation, "Done"
    End If

End Sub
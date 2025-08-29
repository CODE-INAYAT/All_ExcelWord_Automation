' Attendace Mark using T&P UID (Optional Column), Roll NO., Branch, Division
Option Explicit
Sub MarkAttendanceFromExternalFile()
    Dim wsMark As Worksheet
    Dim wbData As Workbook
    Dim wsData As Worksheet
    Dim dataArray As Variant
    Dim dataFilePath As Variant
    Dim missingColsMsg As String
    Dim i As Long

    ' --- STEP 1: Set the "Mark" sheet to the currently active sheet ---
    Set wsMark = ThisWorkbook.ActiveSheet
    MsgBox "The current sheet '" & wsMark.Name & "' will be used as the 'Attendance Mark' sheet.", vbInformation, "Step 1 of 2"

    ' --- STEP 2: Prompt user to select the "Attendance Data" Excel file ---
    dataFilePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx; *.xls; *.xlsm), *.xlsx; *.xls; *.xlsm", _
        Title:="Please select the 'Attendance Data' Excel file")

    ' Handle cancellation
    If dataFilePath = False Then
        MsgBox "Operation cancelled. No file was selected.", vbExclamation, "Cancelled"
        Exit Sub
    End If

    ' --- STEP 3: Open the selected file, read data into an array, and close it ---
    On Error Resume Next
    Application.ScreenUpdating = False ' Prevent screen flicker
    Set wbData = Workbooks.Open(Filename:=dataFilePath, ReadOnly:=True)
    If wbData Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "Error: Could not open the selected file. Please ensure it is a valid Excel file and not password protected.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Find the first visible worksheet in the data workbook
    Set wsData = Nothing
    For i = 1 To wbData.Worksheets.Count
        If wbData.Worksheets(i).Visible = xlSheetVisible Then
            Set wsData = wbData.Worksheets(i)
            Exit For
        End If
    Next i

    If wsData Is Nothing Then
        wbData.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Error: The selected workbook does not contain any visible worksheets.", vbCritical
        Exit Sub
    End If
    
    ' Read all data from the first visible sheet into a variant array for speed
    dataArray = wsData.UsedRange.Value
    
    ' Close the external workbook immediately; we now have the data in memory
    wbData.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Set wbData = Nothing
    Set wsData = Nothing

    ' --- STEP 4: Find column indices and validate mandatory columns ---
    ' Find columns in the "Mark" sheet
    Dim colBranchMark As Long, colDivMark As Long, colRollMark As Long, colUIDMark As Long, colYearMark As Long
    colBranchMark = FindColumn(wsMark, "Branch")
    colDivMark = FindColumn(wsMark, "Division")
    colRollMark = FindColumn(wsMark, "Roll No.")
    colUIDMark = FindColumn(wsMark, "T&P UID")
    colYearMark = FindColumn(wsMark, "Year")

    ' Find columns in the data array (read from the external file)
    Dim colBranchData As Long, colDivData As Long, colRollData As Long, colUIDData As Long
    colBranchData = FindColumnInArray(dataArray, "Branch")
    colDivData = FindColumnInArray(dataArray, "Division")
    colRollData = FindColumnInArray(dataArray, "Roll No.")
    colUIDData = FindColumnInArray(dataArray, "T&P UID")

    ' Check for mandatory columns
    missingColsMsg = ""
    If colBranchMark = 0 Then missingColsMsg = missingColsMsg & vbCrLf & "- Branch (in " & wsMark.Name & " sheet)"
    If colDivMark = 0 Then missingColsMsg = missingColsMsg & vbCrLf & "- Division (in " & wsMark.Name & " sheet)"
    If colRollMark = 0 Then missingColsMsg = missingColsMsg & vbCrLf & "- Roll No. (in " & wsMark.Name & " sheet)"
    If colBranchData = 0 Then missingColsMsg = missingColsMsg & vbCrLf & "- Branch (in the selected data file)"
    If colDivData = 0 Then missingColsMsg = missingColsMsg & vbCrLf & "- Division (in the selected data file)"
    If colRollData = 0 Then missingColsMsg = missingColsMsg & vbCrLf & "- Roll No. (in the selected data file)"

    If missingColsMsg <> "" Then
        MsgBox "Execution stopped. The following mandatory columns are missing:" & missingColsMsg, vbCritical, "Missing Columns"
        Exit Sub
    End If

    ' --- STEP 5: Perform the attendance marking logic ---
    Dim lastRowMark As Long, attendanceCol As Long
    lastRowMark = wsMark.Cells(wsMark.Rows.Count, colBranchMark).End(xlUp).Row
    attendanceCol = FindColumn(wsMark, "Attendance")
    If attendanceCol = 0 Then
        attendanceCol = wsMark.Cells(1, wsMark.Columns.Count).End(xlToLeft).Column + 1
        wsMark.Cells(1, attendanceCol).Value = "Attendance"
    End If

    Dim j As Long, found As Boolean
    Dim uidMark As String, branchMark As String, divMark As String, rollMark As String
    Dim uidData As String, branchData As String, divData As String, rollData As String
    
    For i = 2 To lastRowMark
        branchMark = wsMark.Cells(i, colBranchMark).Value
        divMark = wsMark.Cells(i, colDivMark).Value
        rollMark = wsMark.Cells(i, colRollMark).Value
        found = False

        If IsEmpty(branchMark) Or IsEmpty(divMark) Or IsEmpty(rollMark) Then
            wsMark.Cells(i, attendanceCol).Value = "A"
            GoTo NextMarkRow
        End If
        
        If colUIDMark > 0 Then uidMark = wsMark.Cells(i, colUIDMark).Value Else uidMark = ""

        ' Loop through the data array (from the external file)
        For j = 2 To UBound(dataArray, 1)
            branchData = dataArray(j, colBranchData)
            divData = dataArray(j, colDivData)
            rollData = dataArray(j, colRollData)
            
            If branchMark = branchData And divMark = divData And rollMark = rollData Then
                Dim uidCheckPassed As Boolean: uidCheckPassed = True
                If colUIDMark > 0 And colUIDData > 0 Then
                    uidData = dataArray(j, colUIDData)
                    If uidMark <> uidData Then uidCheckPassed = False
                End If
                If uidCheckPassed Then
                    wsMark.Cells(i, attendanceCol).Value = "P": found = True: Exit For
                End If
            End If
        Next j

        If Not found Then wsMark.Cells(i, attendanceCol).Value = "A"
NextMarkRow:
    Next i

    ' --- STEP 6: Generate the report ---
    Dim wsReport As Worksheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Attendance Report")
    On Error GoTo 0 ' Immediately reset error handling

    ' Corrected block If statement
    If Not wsReport Is Nothing Then
        ' If the sheet exists, clear it
        wsReport.Cells.Clear
    Else
        ' If it doesn't exist, create it
        Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsReport.Name = "Attendance Report"
    End If

    GenerateReport wsMark, wsReport, colBranchMark, colDivMark, colYearMark, attendanceCol, lastRowMark

    MsgBox "Attendance marking and report generation completed!", vbInformation, "Success"
End Sub


' ======================================================================================================
' HELPER FUNCTIONS
' ======================================================================================================

' Finds a column header in a worksheet and returns its column number.
Function FindColumn(ws As Worksheet, header As String) As Long
    Dim lastCol As Long, i As Long
    On Error Resume Next
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If Err.Number <> 0 Then FindColumn = 0: Exit Function
    On Error GoTo 0
    For i = 1 To lastCol
        If Trim(ws.Cells(1, i).Value) = Trim(header) Then
            FindColumn = i: Exit Function
        End If
    Next i
    FindColumn = 0
End Function

' Finds a column header in the first row of a 2D array and returns its column number.
Function FindColumnInArray(dataArr As Variant, header As String) As Long
    Dim i As Long
    If Not IsArray(dataArr) Then FindColumnInArray = 0: Exit Function
    For i = 1 To UBound(dataArr, 2)
        If Trim(dataArr(1, i)) = Trim(header) Then
            FindColumnInArray = i: Exit Function
        End If
    Next i
    FindColumnInArray = 0
End Function

' (The rest of the helper functions are unchanged and work perfectly)

Function SortKeys(keys As Variant) As Variant
    Dim i As Long, j As Long, temp As Variant
    Dim keyArray() As String: ReDim keyArray(0 To UBound(keys))
    For i = 0 To UBound(keys): keyArray(i) = keys(i): Next i
    For i = 0 To UBound(keyArray) - 1
        For j = i + 1 To UBound(keyArray)
            If keyArray(i) > keyArray(j) Then
                temp = keyArray(i): keyArray(i) = keyArray(j): keyArray(j) = temp
            End If
        Next j
    Next i
    SortKeys = keyArray
End Function

Function SortYearKeys(keys As Variant, isMultiLevel As Boolean) As Variant
    Dim i As Long, j As Long, k As Long, l As Long
    Dim temp As Variant, keyArray() As String, yearOrder As Variant, sortedArray() As String
    ReDim keyArray(0 To UBound(keys))
    yearOrder = Array("FE", "SE", "TE", "BE")
    For i = 0 To UBound(keys): keyArray(i) = keys(i): Next i
    ReDim sortedArray(0 To UBound(keys))
    Dim sortedIndex As Long: sortedIndex = 0
    For i = 0 To UBound(yearOrder)
        Dim year As String: year = yearOrder(i)
        Dim tempKeys() As String, tempCount As Long: tempCount = 0
        For j = 0 To UBound(keyArray)
            If InStr(keyArray(j), year & "-") = 1 Or (Not isMultiLevel And keyArray(j) = year) Then
                ReDim Preserve tempKeys(0 To tempCount)
                tempKeys(tempCount) = keyArray(j): tempCount = tempCount + 1
            End If
        Next j
        If tempCount > 0 Then
            For k = 0 To tempCount - 2
                For l = k + 1 To tempCount - 1
                    If tempKeys(k) > tempKeys(l) Then
                        temp = tempKeys(k): tempKeys(k) = tempKeys(l): tempKeys(l) = temp
                    End If
                Next l
            Next k
            For k = 0 To tempCount - 1
                sortedArray(sortedIndex) = tempKeys(k): sortedIndex = sortedIndex + 1
            Next k
        End If
    Next i
    If sortedIndex > 0 Then
        ReDim Preserve sortedArray(0 To sortedIndex - 1): SortYearKeys = sortedArray
    Else
        SortYearKeys = keyArray
    End If
End Function

Sub GenerateReport(wsMark As Worksheet, wsReport As Worksheet, colBranch As Long, colDiv As Long, colYear As Long, colAttendance As Long, lastRow As Long)
    Dim branchDict As Object, branchDivDict As Object, yearDict As Object, yearBranchDict As Object, yearBranchDivDict As Object
    Set branchDict = CreateObject("Scripting.Dictionary"): Set branchDivDict = CreateObject("Scripting.Dictionary")
    Set yearDict = CreateObject("Scripting.Dictionary"): Set yearBranchDict = CreateObject("Scripting.Dictionary")
    Set yearBranchDivDict = CreateObject("Scripting.Dictionary")
    Dim i As Long, totalRegistered As Long, totalAttended As Long, currentRow As Long
    totalRegistered = 0: totalAttended = 0
    For i = 2 To lastRow
        Dim branch As Variant, div As Variant, year As Variant, tempArray As Variant
        branch = wsMark.Cells(i, colBranch).Value: div = wsMark.Cells(i, colDiv).Value
        If IsEmpty(branch) Or IsEmpty(div) Then GoTo NextReportRow
        If Not branchDict.Exists(branch) Then branchDict.Add branch, Array(0, 0)
        tempArray = branchDict(branch): tempArray(0) = tempArray(0) + 1: If wsMark.Cells(i, colAttendance).Value = "P" Then tempArray(1) = tempArray(1) + 1: branchDict(branch) = tempArray
        Dim branchDivKey As String: branchDivKey = branch & "-" & div
        If Not branchDivDict.Exists(branchDivKey) Then branchDivDict.Add branchDivKey, Array(0, 0)
        tempArray = branchDivDict(branchDivKey): tempArray(0) = tempArray(0) + 1: If wsMark.Cells(i, colAttendance).Value = "P" Then tempArray(1) = tempArray(1) + 1: branchDivDict(branchDivKey) = tempArray
        If colYear > 0 Then
            year = wsMark.Cells(i, colYear).Value
            If Not IsEmpty(year) Then
                If Not yearDict.Exists(year) Then yearDict.Add year, Array(0, 0)
                tempArray = yearDict(year): tempArray(0) = tempArray(0) + 1: If wsMark.Cells(i, colAttendance).Value = "P" Then tempArray(1) = tempArray(1) + 1: yearDict(year) = tempArray
                Dim yearBranchKey As String: yearBranchKey = year & "-" & branch
                If Not yearBranchDict.Exists(yearBranchKey) Then yearBranchDict.Add yearBranchKey, Array(0, 0)
                tempArray = yearBranchDict(yearBranchKey): tempArray(0) = tempArray(0) + 1: If wsMark.Cells(i, colAttendance).Value = "P" Then tempArray(1) = tempArray(1) + 1: yearBranchDict(yearBranchKey) = tempArray
                Dim yearBranchDivKey As String: yearBranchDivKey = year & "-" & branch & "-" & div
                If Not yearBranchDivDict.Exists(yearBranchDivKey) Then yearBranchDivDict.Add yearBranchDivKey, Array(0, 0)
                tempArray = yearBranchDivDict(yearBranchDivKey): tempArray(0) = tempArray(0) + 1: If wsMark.Cells(i, colAttendance).Value = "P" Then tempArray(1) = tempArray(1) + 1: yearBranchDivDict(yearBranchDivKey) = tempArray
            End If
        End If
        totalRegistered = totalRegistered + 1: If wsMark.Cells(i, colAttendance).Value = "P" Then totalAttended = totalAttended + 1
NextReportRow:
    Next i
    currentRow = 1
    Dim sortedKeys As Variant, key As Variant, sectionRegTotal As Long, sectionAttTotal As Long
    ' Report by Branch
    wsReport.Cells(currentRow, 1).Value = "Report by Branch": wsReport.Cells(currentRow, 1).Font.Bold = True: currentRow = currentRow + 1
    wsReport.Range(wsReport.Cells(currentRow, 1), wsReport.Cells(currentRow, 4)).Value = Array("Branch", "Total Registered", "Total Attended", "Percentage"): currentRow = currentRow + 1
    sortedKeys = SortKeys(branchDict.Keys): sectionRegTotal = 0: sectionAttTotal = 0
    For Each key In sortedKeys
        wsReport.Cells(currentRow, 1).Value = key: wsReport.Cells(currentRow, 2).Value = branchDict(key)(0): wsReport.Cells(currentRow, 3).Value = branchDict(key)(1)
        sectionRegTotal = sectionRegTotal + branchDict(key)(0): sectionAttTotal = sectionAttTotal + branchDict(key)(1)
        If branchDict(key)(0) > 0 Then wsReport.Cells(currentRow, 4).Value = branchDict(key)(1) / branchDict(key)(0) Else wsReport.Cells(currentRow, 4).Value = 0
        wsReport.Cells(currentRow, 4).NumberFormat = "0.00%": currentRow = currentRow + 1
    Next key
    wsReport.Cells(currentRow, 1).Value = "Total": wsReport.Cells(currentRow, 1).Font.Bold = True: wsReport.Cells(currentRow, 2).Value = sectionRegTotal: wsReport.Cells(currentRow, 3).Value = sectionAttTotal
    If sectionRegTotal > 0 Then wsReport.Cells(currentRow, 4).Value = sectionAttTotal / sectionRegTotal Else wsReport.Cells(currentRow, 4).Value = 0
    wsReport.Cells(currentRow, 4).NumberFormat = "0.00%": currentRow = currentRow + 2
    ' Report by Branch & Division
    wsReport.Cells(currentRow, 1).Value = "Report by Branch & Division": wsReport.Cells(currentRow, 1).Font.Bold = True: currentRow = currentRow + 1
    wsReport.Range(wsReport.Cells(currentRow, 1), wsReport.Cells(currentRow, 5)).Value = Array("Branch", "Division", "Total Registered", "Total Attended", "Percentage"): currentRow = currentRow + 1
    sortedKeys = SortKeys(branchDivDict.Keys): sectionRegTotal = 0: sectionAttTotal = 0
    For Each key In sortedKeys
        Dim parts As Variant: parts = Split(key, "-")
        wsReport.Cells(currentRow, 1).Value = parts(0): wsReport.Cells(currentRow, 2).Value = parts(1): wsReport.Cells(currentRow, 3).Value = branchDivDict(key)(0): wsReport.Cells(currentRow, 4).Value = branchDivDict(key)(1)
        sectionRegTotal = sectionRegTotal + branchDivDict(key)(0): sectionAttTotal = sectionAttTotal + branchDivDict(key)(1)
        If branchDivDict(key)(0) > 0 Then wsReport.Cells(currentRow, 5).Value = branchDivDict(key)(1) / branchDivDict(key)(0) Else wsReport.Cells(currentRow, 5).Value = 0
        wsReport.Cells(currentRow, 5).NumberFormat = "0.00%": currentRow = currentRow + 1
    Next key
    wsReport.Cells(currentRow, 1).Value = "Total": wsReport.Cells(currentRow, 1).Font.Bold = True: wsReport.Cells(currentRow, 3).Value = sectionRegTotal: wsReport.Cells(currentRow, 4).Value = sectionAttTotal
    If sectionRegTotal > 0 Then wsReport.Cells(currentRow, 5).Value = sectionAttTotal / sectionRegTotal Else wsReport.Cells(currentRow, 5).Value = 0
    wsReport.Cells(currentRow, 5).NumberFormat = "0.00%": currentRow = currentRow + 2
    If colYear > 0 And yearDict.Count > 0 Then
        ' Report by Year, Branch, & Division sections...
        wsReport.Cells(currentRow, 1).Value = "Report by Year": wsReport.Cells(currentRow, 1).Font.Bold = True: currentRow = currentRow + 1
        wsReport.Range(wsReport.Cells(currentRow, 1), wsReport.Cells(currentRow, 4)).Value = Array("Year", "Total Registered", "Total Attended", "Percentage"): currentRow = currentRow + 1
        sortedKeys = SortYearKeys(yearDict.Keys, False): sectionRegTotal = 0: sectionAttTotal = 0
        For Each key In sortedKeys
            wsReport.Cells(currentRow, 1).Value = key: wsReport.Cells(currentRow, 2).Value = yearDict(key)(0): wsReport.Cells(currentRow, 3).Value = yearDict(key)(1)
            sectionRegTotal = sectionRegTotal + yearDict(key)(0): sectionAttTotal = sectionAttTotal + yearDict(key)(1)
            If yearDict(key)(0) > 0 Then wsReport.Cells(currentRow, 4).Value = yearDict(key)(1) / yearDict(key)(0) Else wsReport.Cells(currentRow, 4).Value = 0
            wsReport.Cells(currentRow, 4).NumberFormat = "0.00%": currentRow = currentRow + 1
        Next key
        wsReport.Cells(currentRow, 1).Value = "Total": wsReport.Cells(currentRow, 1).Font.Bold = True: wsReport.Cells(currentRow, 2).Value = sectionRegTotal: wsReport.Cells(currentRow, 3).Value = sectionAttTotal
        If sectionRegTotal > 0 Then wsReport.Cells(currentRow, 4).Value = sectionAttTotal / sectionRegTotal Else wsReport.Cells(currentRow, 4).Value = 0
        wsReport.Cells(currentRow, 4).NumberFormat = "0.00%": currentRow = currentRow + 2
        wsReport.Cells(currentRow, 1).Value = "Report by Year & Branch": wsReport.Cells(currentRow, 1).Font.Bold = True: currentRow = currentRow + 1
        wsReport.Range(wsReport.Cells(currentRow, 1), wsReport.Cells(currentRow, 5)).Value = Array("Year", "Branch", "Total Registered", "Total Attended", "Percentage"): currentRow = currentRow + 1
        sortedKeys = SortYearKeys(yearBranchDict.Keys, True): sectionRegTotal = 0: sectionAttTotal = 0
        For Each key In sortedKeys
            parts = Split(key, "-")
            wsReport.Cells(currentRow, 1).Value = parts(0): wsReport.Cells(currentRow, 2).Value = parts(1): wsReport.Cells(currentRow, 3).Value = yearBranchDict(key)(0): wsReport.Cells(currentRow, 4).Value = yearBranchDict(key)(1)
            sectionRegTotal = sectionRegTotal + yearBranchDict(key)(0): sectionAttTotal = sectionAttTotal + yearBranchDict(key)(1)
            If yearBranchDict(key)(0) > 0 Then wsReport.Cells(currentRow, 5).Value = yearBranchDict(key)(1) / yearBranchDict(key)(0) Else wsReport.Cells(currentRow, 5).Value = 0
            wsReport.Cells(currentRow, 5).NumberFormat = "0.00%": currentRow = currentRow + 1
        Next key
        wsReport.Cells(currentRow, 1).Value = "Total": wsReport.Cells(currentRow, 1).Font.Bold = True: wsReport.Cells(currentRow, 3).Value = sectionRegTotal: wsReport.Cells(currentRow, 4).Value = sectionAttTotal
        If sectionRegTotal > 0 Then wsReport.Cells(currentRow, 5).Value = sectionAttTotal / sectionRegTotal Else wsReport.Cells(currentRow, 5).Value = 0
        wsReport.Cells(currentRow, 5).NumberFormat = "0.00%": currentRow = currentRow + 2
        wsReport.Cells(currentRow, 1).Value = "Report by Year, Branch, Division": wsReport.Cells(currentRow, 1).Font.Bold = True: currentRow = currentRow + 1
        wsReport.Range(wsReport.Cells(currentRow, 1), wsReport.Cells(currentRow, 6)).Value = Array("Year", "Branch", "Division", "Total Registered", "Total Attended", "Percentage"): currentRow = currentRow + 1
        sortedKeys = SortYearKeys(yearBranchDivDict.Keys, True): sectionRegTotal = 0: sectionAttTotal = 0
        For Each key In sortedKeys
            parts = Split(key, "-")
            wsReport.Cells(currentRow, 1).Value = parts(0): wsReport.Cells(currentRow, 2).Value = parts(1): wsReport.Cells(currentRow, 3).Value = parts(2)
            wsReport.Cells(currentRow, 4).Value = yearBranchDivDict(key)(0): wsReport.Cells(currentRow, 5).Value = yearBranchDivDict(key)(1)
            sectionRegTotal = sectionRegTotal + yearBranchDivDict(key)(0): sectionAttTotal = sectionAttTotal + yearBranchDivDict(key)(1)
            If yearBranchDivDict(key)(0) > 0 Then wsReport.Cells(currentRow, 6).Value = yearBranchDivDict(key)(1) / yearBranchDivDict(key)(0) Else wsReport.Cells(currentRow, 6).Value = 0
            wsReport.Cells(currentRow, 6).NumberFormat = "0.00%": currentRow = currentRow + 1
        Next key
        wsReport.Cells(currentRow, 1).Value = "Total": wsReport.Cells(currentRow, 1).Font.Bold = True: wsReport.Cells(currentRow, 4).Value = sectionRegTotal: wsReport.Cells(currentRow, 5).Value = sectionAttTotal
        If sectionRegTotal > 0 Then wsReport.Cells(currentRow, 6).Value = sectionAttTotal / sectionRegTotal Else wsReport.Cells(currentRow, 6).Value = 0
        wsReport.Cells(currentRow, 6).NumberFormat = "0.00%": currentRow = currentRow + 2
    End If
    ' Overall Summary
    wsReport.Cells(currentRow, 1).Value = "Overall Summary": wsReport.Cells(currentRow, 1).Font.Bold = True: currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Total Registered": wsReport.Cells(currentRow, 2).Value = totalRegistered: currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Total Attended": wsReport.Cells(currentRow, 2).Value = totalAttended: currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Percentage"
    If totalRegistered > 0 Then wsReport.Cells(currentRow, 2).Value = totalAttended / totalRegistered Else wsReport.Cells(currentRow, 2).Value = 0
    wsReport.Cells(currentRow, 2).NumberFormat = "0.00%"
    wsReport.Columns.AutoFit
End Sub
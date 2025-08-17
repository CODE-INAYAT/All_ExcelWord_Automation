' Compare Two sheets Attendance Mark & Attendance Data, to mark the P/A in Attendace Mark sheet. Also, Generates an Report

Sub MarkAttendanceAndGenerateReport()
    Dim wsMark As Worksheet
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRowMark As Long
    Dim lastRowData As Long
    Dim i As Long
    Dim j As Long
    Dim uidMark As String
    Dim branchMark As String
    Dim divMark As String
    Dim rollMark As String
    Dim uidData As String
    Dim branchData As String
    Dim divData As String
    Dim rollData As String
    Dim found As Boolean
    Dim colUIDMark As Long
    Dim colBranchMark As Long
    Dim colDivMark As Long
    Dim colRollMark As Long
    Dim colUIDData As Long
    Dim colBranchData As Long
    Dim colDivData As Long
    Dim colRollData As Long
    Dim colYearMark As Long
    Dim attendanceCol As Long
    Dim lastColMark As Long
    
    ' Set references to the worksheets
    Set wsMark = ThisWorkbook.Sheets("Attendance Mark")
    Set wsData = ThisWorkbook.Sheets("Attendance Data")
    
    ' Find the last row in both sheets
    lastRowMark = wsMark.Cells(wsMark.Rows.Count, 1).End(xlUp).Row
    lastRowData = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' Find column indices in Attendance Mark sheet
    colUIDMark = FindColumn(wsMark, "T&P UID")
    colBranchMark = FindColumn(wsMark, "Branch")
    colDivMark = FindColumn(wsMark, "Division")
    colRollMark = FindColumn(wsMark, "Roll No.")
    colYearMark = FindColumn(wsMark, "Year")
    
    ' Find column indices in Attendance Data sheet
    colUIDData = FindColumn(wsData, "T&P UID")
    colBranchData = FindColumn(wsData, "Branch")
    colDivData = FindColumn(wsData, "Division")
    colRollData = FindColumn(wsData, "Roll No.")
    
    ' Check if all required columns were found
    If colUIDMark = 0 Or colBranchMark = 0 Or colDivMark = 0 Or colRollMark = 0 Or colYearMark = 0 Or _
       colUIDData = 0 Or colBranchData = 0 Or colDivData = 0 Or colRollData = 0 Then
        MsgBox "One or more required columns (T&P UID, Branch, Division, Roll No., Year) not found in the sheets.", vbCritical
        Exit Sub
    End If
    
    ' Find or create Attendance column in Attendance Mark sheet
    lastColMark = wsMark.Cells(1, wsMark.Columns.Count).End(xlToLeft).Column
    attendanceCol = FindColumn(wsMark, "Attendance")
    If attendanceCol = 0 Then
        attendanceCol = lastColMark + 1
        wsMark.Cells(1, attendanceCol).Value = "Attendance"
    End If
    
    ' Loop through each row in Attendance Mark sheet to mark attendance
    For i = 2 To lastRowMark
        uidMark = wsMark.Cells(i, colUIDMark).Value
        branchMark = wsMark.Cells(i, colBranchMark).Value
        divMark = wsMark.Cells(i, colDivMark).Value
        rollMark = wsMark.Cells(i, colRollMark).Value
        found = False
        
        For j = 2 To lastRowData
            uidData = wsData.Cells(j, colUIDData).Value
            branchData = wsData.Cells(j, colBranchData).Value
            divData = wsData.Cells(j, colDivData).Value
            rollData = wsData.Cells(j, colRollData).Value
            
            If uidMark = uidData And branchMark = branchData And divMark = divData And rollMark = rollData Then
                wsMark.Cells(i, attendanceCol).Value = "P"
                found = True
                Exit For
            End If
        Next j
        
        If Not found Then
            wsMark.Cells(i, attendanceCol).Value = "A"
        End If
    Next i
    
    ' Create or clear Attendance Report sheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Attendance Report")
    If Not wsReport Is Nothing Then
        wsReport.Cells.Clear
    Else
        Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsReport.Name = "Attendance Report"
    End If
    On Error GoTo 0
    
    ' Generate summaries
    GenerateReport wsMark, wsReport, colBranchMark, colDivMark, colYearMark, attendanceCol, lastRowMark
    
    MsgBox "Attendance marking and report generation completed!", vbInformation
End Sub

' Helper function to find column index by header name
Function FindColumn(ws As Worksheet, header As String) As Long
    Dim lastCol As Long
    Dim i As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        If ws.Cells(1, i).Value = header Then
            FindColumn = i
            Exit Function
        End If
    Next i
    FindColumn = 0
End Function

' Helper function to sort an array of strings alphabetically
Function SortKeys(keys As Variant) As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim keyArray() As String
    ReDim keyArray(0 To UBound(keys))
    
    ' Copy keys to array
    For i = 0 To UBound(keys)
        keyArray(i) = keys(i)
    Next i
    
    ' Bubble sort
    For i = 0 To UBound(keyArray) - 1
        For j = i + 1 To UBound(keyArray)
            If keyArray(i) > keyArray(j) Then
                temp = keyArray(i)
                keyArray(i) = keyArray(j)
                keyArray(j) = temp
            End If
        Next j
    Next i
    
    SortKeys = keyArray
End Function

' Helper function to sort year-related keys with custom order (FE, SE, TE, BE)
Function SortYearKeys(keys As Variant, isMultiLevel As Boolean) As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim keyArray() As String
    Dim yearOrder As Variant
    Dim sortedArray() As String
    ReDim keyArray(0 To UBound(keys))
    
    ' Define custom year order
    yearOrder = Array("FE", "SE", "TE", "BE")
    
    ' Copy keys to array
    For i = 0 To UBound(keys)
        keyArray(i) = keys(i)
    Next i
    
    ' Sort based on year order and optionally Branch/Division
    ReDim sortedArray(0 To UBound(keys))
    Dim sortedIndex As Long
    sortedIndex = 0
    
    ' For each year in custom order
    For i = 0 To UBound(yearOrder)
        Dim year As String
        year = yearOrder(i)
        Dim tempKeys() As String
        Dim tempCount As Long
        tempCount = 0
        
        ' Collect keys starting with this year
        For j = 0 To UBound(keyArray)
            If InStr(keyArray(j), year & "-") = 1 Or (Not isMultiLevel And keyArray(j) = year) Then
                ReDim Preserve tempKeys(0 To tempCount)
                tempKeys(tempCount) = keyArray(j)
                tempCount = tempCount + 1
            End If
        Next j
        
        ' Sort tempKeys alphabetically (for Branch or Branch-Division)
        If tempCount > 0 Then
            Dim k As Long, l As Long
            For k = 0 To tempCount - 2
                For l = k + 1 To tempCount - 1
                    If tempKeys(k) > tempKeys(l) Then
                        temp = tempKeys(k)
                        tempKeys(k) = tempKeys(l)
                        tempKeys(l) = temp
                    End If
                Next l
            Next k
            
            ' Add sorted keys to sortedArray
            For k = 0 To tempCount - 1
                sortedArray(sortedIndex) = tempKeys(k)
                sortedIndex = sortedIndex + 1
            Next k
        End If
    Next i
    
    ' Trim sortedArray to actual size
    ReDim Preserve sortedArray(0 To sortedIndex - 1)
    SortYearKeys = sortedArray
End Function

' Function to generate the attendance report
Sub GenerateReport(wsMark As Worksheet, wsReport As Worksheet, colBranch As Long, colDiv As Long, colYear As Long, colAttendance As Long, lastRow As Long)
    Dim i As Long
    Dim branchDict As Object
    Dim branchDivDict As Object
    Dim yearDict As Object
    Dim yearBranchDict As Object
    Dim yearBranchDivDict As Object
    Dim totalRegistered As Long
    Dim totalAttended As Long
    Dim currentRow As Long
    Dim tempArray As Variant
    Dim sortedKeys As Variant
    Dim sectionRegTotal As Long
    Dim sectionAttTotal As Long
    
    ' Initialize dictionaries
    Set branchDict = CreateObject("Scripting.Dictionary")
    Set branchDivDict = CreateObject("Scripting.Dictionary")
    Set yearDict = CreateObject("Scripting.Dictionary")
    Set yearBranchDict = CreateObject("Scripting.Dictionary")
    Set yearBranchDivDict = CreateObject("Scripting.Dictionary")
    totalRegistered = 0
    totalAttended = 0
    
    ' Collect data
    For i = 2 To lastRow
        Dim branch As Variant
        Dim div As Variant
        Dim year As Variant
        Dim branchDivKey As Variant
        Dim yearBranchKey As Variant
        Dim yearBranchDivKey As Variant
        
        branch = wsMark.Cells(i, colBranch).Value
        div = wsMark.Cells(i, colDiv).Value
        year = wsMark.Cells(i, colYear).Value
        branchDivKey = branch & "-" & div
        yearBranchKey = year & "-" & branch
        yearBranchDivKey = year & "-" & branch & "-" & div
        
        ' Skip empty or invalid entries
        If IsEmpty(branch) Or IsEmpty(div) Or IsEmpty(year) Then
            GoTo NextRow
        End If
        
        ' Update branch summary
        If Not branchDict.Exists(branch) Then
            branchDict.Add branch, Array(0, 0) ' [Registered, Attended]
        End If
        tempArray = branchDict(branch)
        tempArray(0) = tempArray(0) + 1
        If wsMark.Cells(i, colAttendance).Value = "P" Then
            tempArray(1) = tempArray(1) + 1
        End If
        branchDict(branch) = tempArray
        
        ' Update branch & division summary
        If Not branchDivDict.Exists(branchDivKey) Then
            branchDivDict.Add branchDivKey, Array(0, 0)
        End If
        tempArray = branchDivDict(branchDivKey)
        tempArray(0) = tempArray(0) + 1
        If wsMark.Cells(i, colAttendance).Value = "P" Then
            tempArray(1) = tempArray(1) + 1
        End If
        branchDivDict(branchDivKey) = tempArray
        
        ' Update year summary
        If Not yearDict.Exists(year) Then
            yearDict.Add year, Array(0, 0)
        End If
        tempArray = yearDict(year)
        tempArray(0) = tempArray(0) + 1
        If wsMark.Cells(i, colAttendance).Value = "P" Then
            tempArray(1) = tempArray(1) + 1
        End If
        yearDict(year) = tempArray
        
        ' Update year & branch summary
        If Not yearBranchDict.Exists(yearBranchKey) Then
            yearBranchDict.Add yearBranchKey, Array(0, 0)
        End If
        tempArray = yearBranchDict(yearBranchKey)
        tempArray(0) = tempArray(0) + 1
        If wsMark.Cells(i, colAttendance).Value = "P" Then
            tempArray(1) = tempArray(1) + 1
        End If
        yearBranchDict(yearBranchKey) = tempArray
        
        ' Update year, branch, division summary
        If Not yearBranchDivDict.Exists(yearBranchDivKey) Then
            yearBranchDivDict.Add yearBranchDivKey, Array(0, 0)
        End If
        tempArray = yearBranchDivDict(yearBranchDivKey)
        tempArray(0) = tempArray(0) + 1
        If wsMark.Cells(i, colAttendance).Value = "P" Then
            tempArray(1) = tempArray(1) + 1
        End If
        yearBranchDivDict(yearBranchDivKey) = tempArray
        
        ' Update overall summary
        totalRegistered = totalRegistered + 1
        If wsMark.Cells(i, colAttendance).Value = "P" Then
            totalAttended = totalAttended + 1
        End If
NextRow:
    Next i
    
    ' Write Report to Attendance Report sheet
    currentRow = 1
    
    ' Report by Branch
    wsReport.Cells(currentRow, 1).Value = "Report by Branch"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Branch"
    wsReport.Cells(currentRow, 2).Value = "Total Registered"
    wsReport.Cells(currentRow, 3).Value = "Total Attended"
    wsReport.Cells(currentRow, 4).Value = "Percentage"
    currentRow = currentRow + 1
    sortedKeys = SortKeys(branchDict.Keys)
    sectionRegTotal = 0
    sectionAttTotal = 0
    Dim branchKey As Variant
    For Each branchKey In sortedKeys
        wsReport.Cells(currentRow, 1).Value = branchKey
        wsReport.Cells(currentRow, 2).Value = branchDict(branchKey)(0)
        wsReport.Cells(currentRow, 3).Value = branchDict(branchKey)(1)
        sectionRegTotal = sectionRegTotal + branchDict(branchKey)(0)
        sectionAttTotal = sectionAttTotal + branchDict(branchKey)(1)
        If branchDict(branchKey)(0) > 0 Then
            wsReport.Cells(currentRow, 4).Value = branchDict(branchKey)(1) / branchDict(branchKey)(0)
            wsReport.Cells(currentRow, 4).NumberFormat = "0.00%"
        Else
            wsReport.Cells(currentRow, 4).Value = 0
            wsReport.Cells(currentRow, 4).NumberFormat = "0.00%"
        End If
        currentRow = currentRow + 1
    Next branchKey
    ' Add Total row
    wsReport.Cells(currentRow, 1).Value = "Total"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    wsReport.Cells(currentRow, 2).Value = sectionRegTotal
    wsReport.Cells(currentRow, 3).Value = sectionAttTotal
    If sectionRegTotal > 0 Then
        wsReport.Cells(currentRow, 4).Value = sectionAttTotal / sectionRegTotal
        wsReport.Cells(currentRow, 4).NumberFormat = "0.00%"
    Else
        wsReport.Cells(currentRow, 4).Value = 0
        wsReport.Cells(currentRow, 4).NumberFormat = "0.00%"
    End If
    currentRow = currentRow + 2
    
    ' Report by Branch & Division
    wsReport.Cells(currentRow, 1).Value = "Report by Branch & Division"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Branch"
    wsReport.Cells(currentRow, 2).Value = "Division"
    wsReport.Cells(currentRow, 3).Value = "Total Registered"
    wsReport.Cells(currentRow, 4).Value = "Total Attended"
    wsReport.Cells(currentRow, 5).Value = "Percentage"
    currentRow = currentRow + 1
    sortedKeys = SortKeys(branchDivDict.Keys)
    sectionRegTotal = 0
    sectionAttTotal = 0
    Dim branchDivKeyLoop As Variant
    For Each branchDivKeyLoop In sortedKeys
        Dim parts As Variant
        parts = Split(branchDivKeyLoop, "-")
        wsReport.Cells(currentRow, 1).Value = parts(0)
        wsReport.Cells(currentRow, 2).Value = parts(1)
        wsReport.Cells(currentRow, 3).Value = branchDivDict(branchDivKeyLoop)(0)
        wsReport.Cells(currentRow, 4).Value = branchDivDict(branchDivKeyLoop)(1)
        sectionRegTotal = sectionRegTotal + branchDivDict(branchDivKeyLoop)(0)
        sectionAttTotal = sectionAttTotal + branchDivDict(branchDivKeyLoop)(1)
        If branchDivDict(branchDivKeyLoop)(0) > 0 Then
            wsReport.Cells(currentRow, 5).Value = branchDivDict(branchDivKeyLoop)(1) / branchDivDict(branchDivKeyLoop)(0)
            wsReport.Cells(currentRow, 5).NumberFormat = "0.00%"
        Else
            wsReport.Cells(currentRow, 5).Value = 0
            wsReport.Cells(currentRow, 5).NumberFormat = "0.00%"
        End If
        currentRow = currentRow + 1
    Next branchDivKeyLoop
    ' Add Total row
    wsReport.Cells(currentRow, 1).Value = "Total"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    wsReport.Cells(currentRow, 3).Value = sectionRegTotal
    wsReport.Cells(currentRow, 4).Value = sectionAttTotal
    If sectionRegTotal > 0 Then
        wsReport.Cells(currentRow, 5).Value = sectionAttTotal / sectionRegTotal
        wsReport.Cells(currentRow, 5).NumberFormat = "0.00%"
    Else
        wsReport.Cells(currentRow, 5).Value = 0
        wsReport.Cells(currentRow, 5).NumberFormat = "0.00%"
    End If
    currentRow = currentRow + 2
    
    ' Report by Year
    wsReport.Cells(currentRow, 1).Value = "Report by Year"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Year"
    wsReport.Cells(currentRow, 2).Value = "Total Registered"
    wsReport.Cells(currentRow, 3).Value = "Total Attended"
    wsReport.Cells(currentRow, 4).Value = "Percentage"
    currentRow = currentRow + 1
    sortedKeys = SortYearKeys(yearDict.Keys, False)
    sectionRegTotal = 0
    sectionAttTotal = 0
    Dim yearKey As Variant
    For Each yearKey In sortedKeys
        wsReport.Cells(currentRow, 1).Value = yearKey
        wsReport.Cells(currentRow, 2).Value = yearDict(yearKey)(0)
        wsReport.Cells(currentRow, 3).Value = yearDict(yearKey)(1)
        sectionRegTotal = sectionRegTotal + yearDict(yearKey)(0)
        sectionAttTotal = sectionAttTotal + yearDict(yearKey)(1)
        If yearDict(yearKey)(0) > 0 Then
            wsReport.Cells(currentRow, 4).Value = yearDict(yearKey)(1) / yearDict(yearKey)(0)
            wsReport.Cells(currentRow, 4).NumberFormat = "0.00%"
        Else
            wsReport.Cells(currentRow, 4).Value = 0
            wsReport.Cells(currentRow, 4).NumberFormat = "0.00%"
        End If
        currentRow = currentRow + 1
    Next yearKey
    ' Add Total row
    wsReport.Cells(currentRow, 1).Value = "Total"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    wsReport.Cells(currentRow, 2).Value = sectionRegTotal
    wsReport.Cells(currentRow, 3).Value = sectionAttTotal
    If sectionRegTotal > 0 Then
        wsReport.Cells(currentRow, 4).Value = sectionAttTotal / sectionRegTotal
        wsReport.Cells(currentRow, 4).NumberFormat = "0.00%"
    Else
        wsReport.Cells(currentRow, 4).Value = 0
        wsReport.Cells(currentRow, 4).NumberFormat = "0.00%"
    End If
    currentRow = currentRow + 2
    
    ' Report by Year & Branch
    wsReport.Cells(currentRow, 1).Value = "Report by Year & Branch"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Year"
    wsReport.Cells(currentRow, 2).Value = "Branch"
    wsReport.Cells(currentRow, 3).Value = "Total Registered"
    wsReport.Cells(currentRow, 4).Value = "Total Attended"
    wsReport.Cells(currentRow, 5).Value = "Percentage"
    currentRow = currentRow + 1
    sortedKeys = SortYearKeys(yearBranchDict.Keys, True)
    sectionRegTotal = 0
    sectionAttTotal = 0
    Dim yearBranchKeyLoop As Variant
    For Each yearBranchKeyLoop In sortedKeys
        Dim ybParts As Variant
        ybParts = Split(yearBranchKeyLoop, "-")
        wsReport.Cells(currentRow, 1).Value = ybParts(0)
        wsReport.Cells(currentRow, 2).Value = ybParts(1)
        wsReport.Cells(currentRow, 3).Value = yearBranchDict(yearBranchKeyLoop)(0)
        wsReport.Cells(currentRow, 4).Value = yearBranchDict(yearBranchKeyLoop)(1)
        sectionRegTotal = sectionRegTotal + yearBranchDict(yearBranchKeyLoop)(0)
        sectionAttTotal = sectionAttTotal + yearBranchDict(yearBranchKeyLoop)(1)
        If yearBranchDict(yearBranchKeyLoop)(0) > 0 Then
            wsReport.Cells(currentRow, 5).Value = yearBranchDict(yearBranchKeyLoop)(1) / yearBranchDict(yearBranchKeyLoop)(0)
            wsReport.Cells(currentRow, 5).NumberFormat = "0.00%"
        Else
            wsReport.Cells(currentRow, 5).Value = 0
            wsReport.Cells(currentRow, 5).NumberFormat = "0.00%"
        End If
        currentRow = currentRow + 1
    Next yearBranchKeyLoop
    ' Add Total row
    wsReport.Cells(currentRow, 1).Value = "Total"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    wsReport.Cells(currentRow, 3).Value = sectionRegTotal
    wsReport.Cells(currentRow, 4).Value = sectionAttTotal
    If sectionRegTotal > 0 Then
        wsReport.Cells(currentRow, 5).Value = sectionAttTotal / sectionRegTotal
        wsReport.Cells(currentRow, 5).NumberFormat = "0.00%"
    Else
        wsReport.Cells(currentRow, 5).Value = 0
        wsReport.Cells(currentRow, 5).NumberFormat = "0.00%"
    End If
    currentRow = currentRow + 2
    
    ' Report by Year, Branch, Division
    wsReport.Cells(currentRow, 1).Value = "Report by Year, Branch, Division"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Year"
    wsReport.Cells(currentRow, 2).Value = "Branch"
    wsReport.Cells(currentRow, 3).Value = "Division"
    wsReport.Cells(currentRow, 4).Value = "Total Registered"
    wsReport.Cells(currentRow, 5).Value = "Total Attended"
    wsReport.Cells(currentRow, 6).Value = "Percentage"
    currentRow = currentRow + 1
    sortedKeys = SortYearKeys(yearBranchDivDict.Keys, True)
    sectionRegTotal = 0
    sectionAttTotal = 0
    Dim yearBranchDivKeyLoop As Variant
    For Each yearBranchDivKeyLoop In sortedKeys
        Dim ybdParts As Variant
        ybdParts = Split(yearBranchDivKeyLoop, "-")
        wsReport.Cells(currentRow, 1).Value = ybdParts(0)
        wsReport.Cells(currentRow, 2).Value = ybdParts(1)
        wsReport.Cells(currentRow, 3).Value = ybdParts(2)
        wsReport.Cells(currentRow, 4).Value = yearBranchDivDict(yearBranchDivKeyLoop)(0)
        wsReport.Cells(currentRow, 5).Value = yearBranchDivDict(yearBranchDivKeyLoop)(1)
        sectionRegTotal = sectionRegTotal + yearBranchDivDict(yearBranchDivKeyLoop)(0)
        sectionAttTotal = sectionAttTotal + yearBranchDivDict(yearBranchDivKeyLoop)(1)
        If yearBranchDivDict(yearBranchDivKeyLoop)(0) > 0 Then
            wsReport.Cells(currentRow, 6).Value = yearBranchDivDict(yearBranchDivKeyLoop)(1) / yearBranchDivDict(yearBranchDivKeyLoop)(0)
            wsReport.Cells(currentRow, 6).NumberFormat = "0.00%"
        Else
            wsReport.Cells(currentRow, 6).Value = 0
            wsReport.Cells(currentRow, 6).NumberFormat = "0.00%"
        End If
        currentRow = currentRow + 1
    Next yearBranchDivKeyLoop
    ' Add Total row
    wsReport.Cells(currentRow, 1).Value = "Total"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    wsReport.Cells(currentRow, 4).Value = sectionRegTotal
    wsReport.Cells(currentRow, 5).Value = sectionAttTotal
    If sectionRegTotal > 0 Then
        wsReport.Cells(currentRow, 6).Value = sectionAttTotal / sectionRegTotal
        wsReport.Cells(currentRow, 6).NumberFormat = "0.00%"
    Else
        wsReport.Cells(currentRow, 6).Value = 0
        wsReport.Cells(currentRow, 6).NumberFormat = "0.00%"
    End If
    currentRow = currentRow + 2
    
    ' Overall Summary
    wsReport.Cells(currentRow, 1).Value = "Overall Summary"
    wsReport.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Total Registered"
    wsReport.Cells(currentRow, 2).Value = totalRegistered
    currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Total Attended"
    wsReport.Cells(currentRow, 2).Value = totalAttended
    currentRow = currentRow + 1
    wsReport.Cells(currentRow, 1).Value = "Percentage"
    If totalRegistered > 0 Then
        wsReport.Cells(currentRow, 2).Value = totalAttended / totalRegistered
        wsReport.Cells(currentRow, 2).NumberFormat = "0.00%"
    Else
        wsReport.Cells(currentRow, 2).Value = 0
        wsReport.Cells(currentRow, 2).NumberFormat = "0.00%"
    End If
    
    ' AutoFit columns for better readability
    wsReport.Columns("A:F").AutoFit
End Sub
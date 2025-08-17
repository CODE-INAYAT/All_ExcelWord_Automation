' ===== Generates a report for Total Admitted, Total Registered, Total Present & Percentage =====
Sub GenerateAttendanceReportFormatted()

    ' ===== SETTINGS =====
    Dim attendanceColName As String: attendanceColName = "Attendance"
    
    ' Define grouping combinations
    Dim groupCombos As Variant
    groupCombos = Array( _
        Array("Branch"), _
        Array("Branch", "Division"), _
        Array("Year"), _
        Array("Year", "Branch") _
    )
    
    ' Define sorting options for each grouping combination
    Dim sortingDict As Object
    Set sortingDict = CreateObject("Scripting.Dictionary")
    
    ' Sorting for Array("Branch") - Alphabetical
    sortingDict("Branch") = Array("ALPHABETICAL")
    
    ' Sorting for Array("Branch", "Division") - Branch: Alphabetical, Division: Alphabetical
    sortingDict("Branch-Division") = Array("ALPHABETICAL", "ALPHABETICAL")
    
    ' Sorting for Array("Year") - Custom order
    sortingDict("Year") = Array("FE,SE,TE,BE")
    
    ' Sorting for Array("Year", "Branch") - Year: Custom, Branch: Alphabetical
    sortingDict("Year-Branch") = Array("FE,SE,TE,BE", "ALPHABETICAL")
    
    ' Define Total Admitted values for each group
    Dim totalAdmittedDict As Object
    Set totalAdmittedDict = CreateObject("Scripting.Dictionary")
    
    ' Branch totals
    totalAdmittedDict("AI&DS") = 139
    totalAdmittedDict("AI&ML") = 70
    totalAdmittedDict("CIVIL") = 69
    totalAdmittedDict("COMP") = 210
    totalAdmittedDict("CS&E") = 70
    totalAdmittedDict("E&CS") = 70
    totalAdmittedDict("E&TC") = 142
    totalAdmittedDict("IOT") = 36
    totalAdmittedDict("IT") = 210
    totalAdmittedDict("M&ME") = 71
    totalAdmittedDict("MECH") = 73
    
    ' Branch-Division totals
    totalAdmittedDict("AI&DS-A") = 70
    totalAdmittedDict("AI&DS-B") = 69
    totalAdmittedDict("AI&ML-NA") = 70
    totalAdmittedDict("CIVIL-NA") = 69
    totalAdmittedDict("COMP-A") = 69
    totalAdmittedDict("COMP-B") = 70
    totalAdmittedDict("COMP-C") = 71
    totalAdmittedDict("CS&E-NA") = 70
    totalAdmittedDict("E&CS-NA") = 70
    totalAdmittedDict("E&TC-A") = 70
    totalAdmittedDict("E&TC-B") = 72
    totalAdmittedDict("IOT-NA") = 36
    totalAdmittedDict("IT-A") = 71
    totalAdmittedDict("IT-B") = 70
    totalAdmittedDict("IT-C") = 69
    totalAdmittedDict("M&ME-NA") = 71
    totalAdmittedDict("MECH-NA") = 73
    ' ====================

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Map column names to indexes
    Dim colIndex As Object
    Set colIndex = CreateObject("Scripting.Dictionary")

    Dim i As Integer
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        colIndex(ws.Cells(1, i).Value) = i
    Next i

    If Not colIndex.exists(attendanceColName) Then
        MsgBox "Attendance column not found!"
        Exit Sub
    End If

    ' Create report sheet
    Dim reportWS As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Attendance Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set reportWS = Worksheets.Add
    reportWS.Name = "Attendance Report"

    Dim outputRow As Long: outputRow = 1
    Dim groupCols As Variant
    Dim row As Long, k As Variant
    Dim reportDict As Object
    Dim key As String, attendanceVal As String
    Dim groupTotalAdmitted As Long, groupPresent As Long, groupRegistered As Long
    Dim colHeader As String
    Dim tempArray As Variant
    Dim predefinedTotalAdmitted As Long
    Dim dataStartRow As Long, groupStartRow As Long
    Dim wsName As String
    
    wsName = ws.Name

    For Each groupCols In groupCombos
        Set reportDict = CreateObject("Scripting.Dictionary")
        
        ' Build key-value pairs - now tracking [present, registered]
        For row = 2 To lastRow
            key = ""
            For i = LBound(groupCols) To UBound(groupCols)
                If Not colIndex.exists(groupCols(i)) Then
                    MsgBox "Column '" & groupCols(i) & "' not found in sheet."
                    Exit Sub
                End If
                key = key & ws.Cells(row, colIndex(groupCols(i))).Value & "-"
            Next i
            If Len(key) > 0 Then key = Left(key, Len(key) - 1)

            attendanceVal = ws.Cells(row, colIndex(attendanceColName)).Value

            If Not reportDict.exists(key) Then
                reportDict(key) = Array(0, 0) ' [total present, total registered]
            End If
            
            ' Extract array, modify, then store back
            tempArray = reportDict(key)
            If UCase(attendanceVal) = "P" Then
                tempArray(0) = tempArray(0) + 1 ' total present count
            End If
            tempArray(1) = tempArray(1) + 1 ' total registered count (count all records)
            reportDict(key) = tempArray
        Next row

        ' Sort the keys according to the defined sorting rules
        Dim sortedKeys As Variant
        sortedKeys = SortKeys(reportDict.keys, groupCols, sortingDict)

        ' Report header
        reportWS.Cells(outputRow, 1).Value = "Report by " & Join(groupCols, "-")
        reportWS.Cells(outputRow, 1).Font.Bold = True
        outputRow = outputRow + 1

        colHeader = Join(groupCols, "-")

        ' Column headers
        reportWS.Cells(outputRow, 1).Value = colHeader
        reportWS.Cells(outputRow, 2).Value = "Total Admitted"
        reportWS.Cells(outputRow, 3).Value = "Total Present"
        reportWS.Cells(outputRow, 4).Value = "Total Registered"
        reportWS.Cells(outputRow, 5).Value = "Percentage (Total Registered & Present)"
        reportWS.Cells(outputRow, 6).Value = "Percentage (Total Admitted, Total Registered, Present)"
        reportWS.Range(reportWS.Cells(outputRow, 1), reportWS.Cells(outputRow, 6)).Font.Bold = True
        
        ' Add background color to indicate editable columns
        reportWS.Cells(outputRow, 2).Interior.Color = RGB(255, 255, 200) ' Light yellow for Total Admitted
        reportWS.Cells(outputRow, 3).Interior.Color = RGB(200, 255, 200) ' Light green for Total Present
        reportWS.Cells(outputRow, 4).Interior.Color = RGB(200, 200, 255) ' Light blue for Total Registered
        reportWS.Cells(outputRow, 5).Interior.Color = RGB(240, 240, 240) ' Light gray for calculated percentages
        reportWS.Cells(outputRow, 6).Interior.Color = RGB(240, 240, 240) ' Light gray for calculated percentages
        
        outputRow = outputRow + 1

        dataStartRow = outputRow
        groupTotalAdmitted = 0
        groupPresent = 0
        groupRegistered = 0

        ' Use sorted keys instead of unsorted keys
        For Each k In sortedKeys
            tempArray = reportDict(k)
            
            ' Get predefined Total Admitted value
            If totalAdmittedDict.exists(k) Then
                predefinedTotalAdmitted = totalAdmittedDict(k)
            Else
                predefinedTotalAdmitted = 0
                MsgBox "Warning: No predefined Total Admitted value found for group '" & k & "'. Using 0."
            End If
            
            reportWS.Cells(outputRow, 1).Value = k
            
            ' All three main columns are now editable values (not formulas)
            reportWS.Cells(outputRow, 2).Value = predefinedTotalAdmitted ' Total Admitted (editable)
            reportWS.Cells(outputRow, 3).Value = tempArray(0) ' Total Present (editable)
            reportWS.Cells(outputRow, 4).Value = tempArray(1) ' Total Registered (editable)
            
            ' Add background colors to indicate editable columns
            reportWS.Cells(outputRow, 2).Interior.Color = RGB(255, 255, 200) ' Light yellow for Total Admitted
            reportWS.Cells(outputRow, 3).Interior.Color = RGB(200, 255, 200) ' Light green for Total Present
            reportWS.Cells(outputRow, 4).Interior.Color = RGB(200, 200, 255) ' Light blue for Total Registered
            
            ' Formula for Percentage (Total Registered & Present) - Present out of Registered
            reportWS.Cells(outputRow, 5).Formula = "=IF(" & reportWS.Cells(outputRow, 4).Address & ">0," & reportWS.Cells(outputRow, 3).Address & "/" & reportWS.Cells(outputRow, 4).Address & ",""N/A"")"
            reportWS.Cells(outputRow, 5).NumberFormat = "0.00%"
            reportWS.Cells(outputRow, 5).Interior.Color = RGB(240, 240, 240) ' Light gray for calculated
            
            ' Formula for Percentage (Total Admitted, Total Registered, Present) - Present out of Admitted
            reportWS.Cells(outputRow, 6).Formula = "=IF(" & reportWS.Cells(outputRow, 2).Address & ">0," & reportWS.Cells(outputRow, 3).Address & "/" & reportWS.Cells(outputRow, 2).Address & ",""N/A"")"
            reportWS.Cells(outputRow, 6).NumberFormat = "0.00%"
            reportWS.Cells(outputRow, 6).Interior.Color = RGB(240, 240, 240) ' Light gray for calculated
            
            outputRow = outputRow + 1
        Next k

        groupStartRow = dataStartRow
        Dim groupEndRow As Long: groupEndRow = outputRow - 1

        ' Group Total Row with formulas that sum the editable values above
        reportWS.Cells(outputRow, 1).Value = "Total"
        reportWS.Cells(outputRow, 2).Formula = "=SUM(" & reportWS.Range(reportWS.Cells(groupStartRow, 2), reportWS.Cells(groupEndRow, 2)).Address & ")"
        reportWS.Cells(outputRow, 3).Formula = "=SUM(" & reportWS.Range(reportWS.Cells(groupStartRow, 3), reportWS.Cells(groupEndRow, 3)).Address & ")"
        reportWS.Cells(outputRow, 4).Formula = "=SUM(" & reportWS.Range(reportWS.Cells(groupStartRow, 4), reportWS.Cells(groupEndRow, 4)).Address & ")"
        reportWS.Cells(outputRow, 5).Formula = "=IF(" & reportWS.Cells(outputRow, 4).Address & ">0," & reportWS.Cells(outputRow, 3).Address & "/" & reportWS.Cells(outputRow, 4).Address & ",""N/A"")"
        reportWS.Cells(outputRow, 5).NumberFormat = "0.00%"
        reportWS.Cells(outputRow, 6).Formula = "=IF(" & reportWS.Cells(outputRow, 2).Address & ">0," & reportWS.Cells(outputRow, 3).Address & "/" & reportWS.Cells(outputRow, 2).Address & ",""N/A"")"
        reportWS.Cells(outputRow, 6).NumberFormat = "0.00%"
        
        reportWS.Range(reportWS.Cells(outputRow, 1), reportWS.Cells(outputRow, 6)).Font.Bold = True
        reportWS.Range(reportWS.Cells(outputRow, 2), reportWS.Cells(outputRow, 6)).Interior.Color = RGB(220, 220, 220) ' Darker gray for totals
        outputRow = outputRow + 2
        
        ' Store first group total row for grand total calculation
        If groupCols(0) = "Branch" And UBound(groupCols) = 0 Then
            Dim grandTotalRowBranch As Long: grandTotalRowBranch = outputRow - 2
        End If
    Next groupCols

    ' Grand Total Section with formulas
    reportWS.Cells(outputRow, 1).Value = "Grand Total"
    reportWS.Cells(outputRow, 2).Formula = "=" & reportWS.Cells(grandTotalRowBranch, 2).Address
    reportWS.Cells(outputRow, 3).Formula = "=" & reportWS.Cells(grandTotalRowBranch, 3).Address
    reportWS.Cells(outputRow, 4).Formula = "=" & reportWS.Cells(grandTotalRowBranch, 4).Address
    reportWS.Cells(outputRow, 5).Formula = "=IF(" & reportWS.Cells(outputRow, 4).Address & ">0," & reportWS.Cells(outputRow, 3).Address & "/" & reportWS.Cells(outputRow, 4).Address & ",""N/A"")"
    reportWS.Cells(outputRow, 5).NumberFormat = "0.00%"
        reportWS.Cells(outputRow, 6).Formula = "=IF(" & reportWS.Cells(outputRow, 2).Address & ">0," & reportWS.Cells(outputRow, 3).Address & "/" & reportWS.Cells(outputRow, 2).Address & ",""N/A"")"
    reportWS.Cells(outputRow, 6).NumberFormat = "0.00%"
    
    reportWS.Range(reportWS.Cells(outputRow, 1), reportWS.Cells(outputRow, 6)).Font.Bold = True
    reportWS.Range(reportWS.Cells(outputRow, 2), reportWS.Cells(outputRow, 6)).Interior.Color = RGB(200, 200, 200) ' Dark gray for grand total

    ' Add a legend/instruction section
    outputRow = outputRow + 3
    reportWS.Cells(outputRow, 1).Value = "INSTRUCTIONS:"
    reportWS.Cells(outputRow, 1).Font.Bold = True
    outputRow = outputRow + 1
    reportWS.Cells(outputRow, 1).Value = "• Yellow cells (Total Admitted) - Editable"
    reportWS.Cells(outputRow + 1, 1).Value = "• Green cells (Total Present) - Editable"
    reportWS.Cells(outputRow + 2, 1).Value = "• Blue cells (Total Registered) - Editable"
    reportWS.Cells(outputRow + 3, 1).Value = "• Gray cells (Percentages & Totals) - Auto-calculated"
    reportWS.Cells(outputRow + 4, 1).Value = "• Change any editable value and percentages will update automatically"
    reportWS.Cells(outputRow + 5, 1).Value = "• Groups are sorted according to predefined rules (see settings)"

    ' Autofit columns
    reportWS.Columns("A:F").AutoFit
    
    ' Enable automatic calculation
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Fully Editable Smart Attendance Report with Custom Sorting generated successfully!" & vbCrLf & vbCrLf & _
           "Features:" & vbCrLf & _
           "• Total Admitted (Yellow) - Editable" & vbCrLf & _
           "• Total Present (Green) - Editable" & vbCrLf & _
           "• Total Registered (Blue) - Editable" & vbCrLf & _
           "• All percentages and totals update automatically" & vbCrLf & _
           "• Custom sorting applied per group settings" & vbCrLf & _
           "• Color-coded for easy identification"

End Sub

' ===== Helper function to sort keys according to defined rules =====
Function SortKeys(keys As Variant, groupCols As Variant, sortingDict As Object) As Variant
    
    Dim groupKey As String
    groupKey = Join(groupCols, "-")
    
    Dim sortRules As Variant
    If sortingDict.exists(groupKey) Then
        sortRules = sortingDict(groupKey)
    Else
        ' Default to alphabetical if no rule defined
        ReDim sortRules(UBound(groupCols))
        Dim j As Integer
        For j = 0 To UBound(groupCols)
            sortRules(j) = "ALPHABETICAL"
        Next j
    End If
    
    ' Convert keys to array for sorting
    Dim keyArray() As String
    ReDim keyArray(UBound(keys))
    Dim i As Integer
    For i = 0 To UBound(keys)
        keyArray(i) = CStr(keys(i))
    Next i
    
    ' Apply multi-level sorting
    Call MultiLevelSort(keyArray, sortRules)
    
    SortKeys = keyArray
End Function

' ===== Helper function for multi-level sorting =====
Sub MultiLevelSort(ByRef keyArray() As String, ByVal sortRules As Variant)
    
    Dim i As Long, j As Long
    Dim temp As String
    Dim swapped As Boolean
    
    ' Bubble sort with custom comparison
    For i = 0 To UBound(keyArray) - 1
        swapped = False
        For j = 0 To UBound(keyArray) - 1 - i
            If CompareKeys(keyArray(j), keyArray(j + 1), sortRules) > 0 Then
                temp = keyArray(j)
                keyArray(j) = keyArray(j + 1)
                keyArray(j + 1) = temp
                swapped = True
            End If
        Next j
        If Not swapped Then Exit For
    Next i
    
End Sub

' ===== Helper function to compare two keys according to sorting rules =====
Function CompareKeys(ByVal key1 As String, ByVal key2 As String, ByVal sortRules As Variant) As Integer
    
    Dim parts1() As String, parts2() As String
    parts1 = Split(key1, "-")
    parts2 = Split(key2, "-")
    
    Dim i As Integer
    For i = 0 To UBound(sortRules)
        If i <= UBound(parts1) And i <= UBound(parts2) Then
            Dim comparison As Integer
            comparison = CompareParts(parts1(i), parts2(i), CStr(sortRules(i)))
            If comparison <> 0 Then
                CompareKeys = comparison
                Exit Function
            End If
        End If
    Next i
    
    CompareKeys = 0 ' Equal
End Function

' ===== Helper function to compare individual parts =====
Function CompareParts(ByVal part1 As String, ByVal part2 As String, ByVal sortRule As String) As Integer
    
    If UCase(sortRule) = "ALPHABETICAL" Then
        ' Alphabetical comparison
        If part1 < part2 Then
            CompareParts = -1
        ElseIf part1 > part2 Then
            CompareParts = 1
        Else
            CompareParts = 0
        End If
    Else
        ' Custom order comparison
        Dim customOrder() As String
        customOrder = Split(sortRule, ",")
        
        Dim pos1 As Integer, pos2 As Integer
        pos1 = GetPositionInArray(part1, customOrder)
        pos2 = GetPositionInArray(part2, customOrder)
        
        If pos1 = -1 And pos2 = -1 Then
            ' Both not found, use alphabetical
            If part1 < part2 Then
                CompareParts = -1
            ElseIf part1 > part2 Then
                CompareParts = 1
            Else
                CompareParts = 0
            End If
        ElseIf pos1 = -1 Then
            ' part1 not found, part2 found - part2 comes first
            CompareParts = 1
        ElseIf pos2 = -1 Then
            ' part1 found, part2 not found - part1 comes first
            CompareParts = -1
        Else
            ' Both found, compare positions
            If pos1 < pos2 Then
                CompareParts = -1
            ElseIf pos1 > pos2 Then
                CompareParts = 1
            Else
                CompareParts = 0
            End If
        End If
    End If
    
End Function

' ===== Helper function to get position in array =====
Function GetPositionInArray(ByVal item As String, ByRef arr() As String) As Integer
    
    Dim i As Integer
    For i = 0 To UBound(arr)
        If UCase(Trim(arr(i))) = UCase(Trim(item)) Then
            GetPositionInArray = i
            Exit Function
        End If
    Next i
    
    GetPositionInArray = -1 ' Not found
End Function
Sub CountDataByColumns()
    Dim ws As Worksheet, newWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim selectedCols As String
    Dim colArray() As String
    Dim i As Long, j As Long
    Dim dictKey As String
    Dim countDict As Object
    Dim outputRow As Long
    Dim colIndices() As Long
    Dim tempArray As Variant
    Dim sheetName As String
    Dim rowSpacing As Long
    Dim spacingCols As String
    Dim spacingColArray() As String
    Dim spacingColIndices() As Long
    Dim lastValue As String
    Dim currentValue As String
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Check if there's data in the worksheet
    If ws.UsedRange.Cells.Count = 1 And IsEmpty(ws.Cells(1, 1)) Then
        MsgBox "No data found in the active sheet!", vbExclamation
        Exit Sub
    End If
    
    ' Find last row and column with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Display available columns to user
    Dim availableCols As String
    availableCols = "Available columns:" & vbNewLine & vbNewLine
    For i = 1 To lastCol
        If Not IsEmpty(ws.Cells(1, i)) Then
            availableCols = availableCols & i & ". " & ws.Cells(1, i).Value & vbNewLine
        End If
    Next i
    
    ' Get column names from user
    selectedCols = InputBox(availableCols & vbNewLine & _
                           "Enter column numbers or names to group by (separated by commas):" & vbNewLine & _
                           "Example: 1,3 or Name,Category", "Select Columns for Counting")
    
    ' Exit if user cancels
    If selectedCols = "" Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Ask for row spacing
    Dim spacingInput As String
    spacingInput = InputBox("Enter number of blank rows to insert between unique groups:" & vbNewLine & _
                           "(Enter 0 for no spacing)", "Row Spacing", "0")
    
    If spacingInput = "" Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    If Not IsNumeric(spacingInput) Or Val(spacingInput) < 0 Then
        MsgBox "Invalid row spacing value. Using 0.", vbExclamation
        rowSpacing = 0
    Else
        rowSpacing = CLng(spacingInput)
    End If
    
    ' If row spacing is requested, ask which columns to use for grouping
    If rowSpacing > 0 Then
        spacingCols = InputBox(availableCols & vbNewLine & _
                              "Enter column numbers or names to determine spacing groups:" & vbNewLine & _
                              "(These columns will be used to identify when to insert blank rows)", _
                              "Spacing Columns", selectedCols)
        
        If spacingCols = "" Then
            MsgBox "Using counting columns for spacing.", vbInformation
            spacingCols = selectedCols
        End If
    End If
    
    ' Parse the selected columns
    colArray = Split(selectedCols, ",")
    ReDim colIndices(UBound(colArray))
    
    ' Convert column names/numbers to indices
    For i = 0 To UBound(colArray)
        colArray(i) = Trim(colArray(i))
        
        ' Check if input is a number
        If IsNumeric(colArray(i)) Then
            colIndices(i) = CLng(colArray(i))
        Else
            ' Find column by name
            Dim found As Boolean
            found = False
            For j = 1 To lastCol
                If UCase(Trim(ws.Cells(1, j).Value)) = UCase(colArray(i)) Then
                    colIndices(i) = j
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                MsgBox "Column '" & colArray(i) & "' not found!", vbExclamation
                Exit Sub
            End If
        End If
        
        ' Validate column index
        If colIndices(i) < 1 Or colIndices(i) > lastCol Then
            MsgBox "Invalid column index: " & colIndices(i), vbExclamation
            Exit Sub
        End If
    Next i
    
    ' Parse spacing columns if needed
    If rowSpacing > 0 Then
        Dim spacingArray() As String
        spacingArray = Split(spacingCols, ",")
        ReDim spacingColIndices(UBound(spacingArray))
        
        For i = 0 To UBound(spacingArray)
            spacingArray(i) = Trim(spacingArray(i))
            
            If IsNumeric(spacingArray(i)) Then
                spacingColIndices(i) = CLng(spacingArray(i))
            Else
                found = False
                For j = 1 To lastCol
                    If UCase(Trim(ws.Cells(1, j).Value)) = UCase(spacingArray(i)) Then
                        spacingColIndices(i) = j
                        found = True
                        Exit For
                    End If
                Next j
                If Not found Then
                    MsgBox "Spacing column '" & spacingArray(i) & "' not found!", vbExclamation
                    Exit Sub
                End If
            End If
            
            If spacingColIndices(i) < 1 Or spacingColIndices(i) > lastCol Then
                MsgBox "Invalid spacing column index: " & spacingColIndices(i), vbExclamation
                Exit Sub
            End If
        Next i
    End If
    
    ' Create dictionary for counting
    Set countDict = CreateObject("Scripting.Dictionary")
    
    ' Count occurrences
    Application.ScreenUpdating = False
    Application.StatusBar = "Processing data..."
    
    For i = 2 To lastRow ' Start from row 2 (assuming row 1 has headers)
        dictKey = ""
        For j = 0 To UBound(colIndices)
            If j > 0 Then dictKey = dictKey & "|"
            dictKey = dictKey & ws.Cells(i, colIndices(j)).Value
        Next j
        
        If countDict.Exists(dictKey) Then
            countDict(dictKey) = countDict(dictKey) + 1
        Else
            countDict(dictKey) = 1
        End If
    Next i
    
    ' Create new worksheet for results
    sheetName = "Count_Summary_" & Format(Now, "yyyymmdd_hhmmss")
    Set newWs = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    newWs.Name = sheetName
    
    ' Add headers to new sheet
    outputRow = 1
    For i = 0 To UBound(colIndices)
        newWs.Cells(outputRow, i + 1).Value = ws.Cells(1, colIndices(i)).Value
    Next i
    newWs.Cells(outputRow, UBound(colIndices) + 2).Value = "Count"
    
    ' Format headers
    With newWs.Range(newWs.Cells(1, 1), newWs.Cells(1, UBound(colIndices) + 2))
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(70, 130, 180) ' Steel Blue
        .Font.Color = RGB(255, 255, 255) ' White
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 25
    End With
    
    ' Sort the dictionary keys before outputting (for proper spacing)
    Dim sortedKeys() As Variant
    Dim keyCount As Long
    keyCount = countDict.Count
    ReDim sortedKeys(keyCount - 1)
    
    i = 0
    Dim key As Variant
    For Each key In countDict.Keys
        sortedKeys(i) = key
        i = i + 1
    Next key
    
    ' Simple bubble sort for the keys
    Dim temp As Variant
    For i = 0 To keyCount - 2
        For j = i + 1 To keyCount - 1
            If sortedKeys(i) > sortedKeys(j) Then
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i
    
    ' Add data to new sheet with spacing
    outputRow = 2
    lastValue = ""
    
    For i = 0 To keyCount - 1
        key = sortedKeys(i)
        Dim splitKey As Variant
        splitKey = Split(key, "|")
        
        ' Check if we need to add spacing
        If rowSpacing > 0 And outputRow > 2 Then
            currentValue = ""
            If rowSpacing > 0 Then
                ' Build current value from spacing columns
                For j = 0 To UBound(spacingColIndices)
                    ' Find the corresponding value in splitKey
                    Dim colIndex As Long
                    For k = 0 To UBound(colIndices)
                        If colIndices(k) = spacingColIndices(j) Then
                            If j > 0 Then currentValue = currentValue & "|"
                            currentValue = currentValue & splitKey(k)
                            Exit For
                        End If
                    Next k
                Next j
                
                ' Add spacing if group changed
                If currentValue <> lastValue Then
                    outputRow = outputRow + rowSpacing
                End If
                lastValue = currentValue
            End If
        ElseIf outputRow = 2 And rowSpacing > 0 Then
            ' Initialize lastValue for first row
            lastValue = ""
            For j = 0 To UBound(spacingColIndices)
                For k = 0 To UBound(colIndices)
                    If colIndices(k) = spacingColIndices(j) Then
                        If j > 0 Then lastValue = lastValue & "|"
                        lastValue = lastValue & splitKey(k)
                        Exit For
                    End If
                Next k
            Next j
        End If
        
        ' Add the data
        For j = 0 To UBound(splitKey)
            newWs.Cells(outputRow, j + 1).Value = splitKey(j)
        Next j
        newWs.Cells(outputRow, UBound(splitKey) + 2).Value = countDict(key)
        outputRow = outputRow + 1
    Next i
    
    ' Format data area (without alternating colors)
    With newWs.Range(newWs.Cells(2, 1), newWs.Cells(outputRow - 1, UBound(colIndices) + 2))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(200, 200, 200)
        .RowHeight = 20
    End With
    
    ' Format count column
    With newWs.Range(newWs.Cells(2, UBound(colIndices) + 2), newWs.Cells(outputRow - 1, UBound(colIndices) + 2))
        .Font.Bold = True
        .Font.Color = RGB(0, 100, 0) ' Dark Green
        .HorizontalAlignment = xlCenter
    End With
    
    ' AutoFit columns
    newWs.Columns.AutoFit
    
    ' Add a summary section
    Dim summaryRow As Long
    summaryRow = outputRow + 2
    
    newWs.Cells(summaryRow, 1).Value = "Summary Statistics"
    newWs.Cells(summaryRow, 1).Font.Bold = True
    newWs.Cells(summaryRow, 1).Font.Size = 14
    
    summaryRow = summaryRow + 1
    newWs.Cells(summaryRow, 1).Value = "Total Unique Combinations:"
    newWs.Cells(summaryRow, 2).Value = countDict.Count
    newWs.Cells(summaryRow, 2).Font.Bold = True
    
    summaryRow = summaryRow + 1
    newWs.Cells(summaryRow, 1).Value = "Total Records Counted:"
    newWs.Cells(summaryRow, 2).Value = Application.WorksheetFunction.Sum(newWs.Range(newWs.Cells(2, UBound(colIndices) + 2), _
                                                                         newWs.Cells(outputRow - 1, UBound(colIndices) + 2)))
    newWs.Cells(summaryRow, 2).Font.Bold = True
    
    summaryRow = summaryRow + 1
    newWs.Cells(summaryRow, 1).Value = "Date Created:"
    newWs.Cells(summaryRow, 2).Value = Format(Now, "dd-mmm-yyyy hh:mm:ss")
    
    If rowSpacing > 0 Then
        summaryRow = summaryRow + 1
        newWs.Cells(summaryRow, 1).Value = "Row Spacing:"
        newWs.Cells(summaryRow, 2).Value = rowSpacing & " rows"
    End If
    
    ' Add borders to summary section
    With newWs.Range(newWs.Cells(outputRow + 2, 1), newWs.Cells(summaryRow, 2))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Add autofilter
    newWs.Range(newWs.Cells(1, 1), newWs.Cells(1, UBound(colIndices) + 2)).AutoFilter
    
    ' Freeze panes (freeze first row)
    newWs.Activate
    newWs.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    newWs.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    MsgBox "Count summary created successfully in sheet: " & sheetName & vbNewLine & _
           "Total unique combinations: " & countDict.Count & vbNewLine & _
           IIf(rowSpacing > 0, "Row spacing: " & rowSpacing & " rows", ""), vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Sub
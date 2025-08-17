' Function to split the sheet based on unique combinations from one or more selected columns,
' with advanced custom sorting and improved sheet naming.
Sub SplitDataByMultipleColumns_Advanced()

    ' --- I. PRE-CHECKS & SETUP ---
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This operation can only be run on a worksheet.", vbExclamation
        Exit Sub
    End If

    Dim sourceWS As Worksheet, tempWS As Worksheet, newWS As Worksheet
    Set sourceWS = ThisWorkbook.ActiveSheet

    ' --- II. OPTIMIZATION ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' --- III. VARIABLE DECLARATION ---
    Dim dataRange As Range, cell As Range
    Dim uniqueCombinations As Object ' Scripting.Dictionary
    Dim key As Variant, r As Long, i As Long
    Dim columnHeaders As String, headerArray() As String, colNums() As Long
    Dim compositeKey As String, splitValues() As String, sanitizedValue As String
    Dim sortedKeys As Variant
    Dim yearColIndexInSplit As Long ' To track the position of the 'Year' column for custom sorting

    On Error GoTo Cleanup ' Ensure all cleanup steps are attempted

    ' --- IV. CREATE A SAFE, TEMPORARY WORKSPACE ---
    Set tempWS = ThisWorkbook.Sheets.Add(After:=sourceWS)
    tempWS.Visible = xlSheetVeryHidden
    sourceWS.Cells.Copy tempWS.Range("A1")

    ' --- V. GET USER INPUT & VALIDATE COLUMNS ---
    Set dataRange = tempWS.UsedRange
    If dataRange.Rows.Count <= 1 Then GoTo Cleanup
    
    columnHeaders = InputBox("Enter the column headers to split by, separated by commas:", _
                             "Split by Multiple Columns", "Branch, Year")
    If columnHeaders = "" Then GoTo Cleanup
    
    headerArray = Split(columnHeaders, ",")
    ReDim colNums(UBound(headerArray))
    yearColIndexInSplit = -1 ' Default to -1 (not found)
    
    For i = 0 To UBound(headerArray)
        Dim currentHeader As String
        currentHeader = Trim(headerArray(i))
        
        On Error Resume Next
        colNums(i) = Application.Match(currentHeader, dataRange.Rows(1), 0)
        On Error GoTo 0
        
        If colNums(i) = 0 Then
            MsgBox "Column header '" & currentHeader & "' not found.", vbExclamation
            GoTo Cleanup
        End If
        
        ' Check if this is the 'Year' column for our custom sort logic
        If LCase(currentHeader) = "year" Then yearColIndexInSplit = i
    Next i

    ' --- VI. GATHER UNIQUE COMBINATIONS ---
    Set uniqueCombinations = CreateObject("Scripting.Dictionary")
    
    For r = 2 To dataRange.Rows.Count
        compositeKey = ""
        For i = 0 To UBound(colNums)
            compositeKey = compositeKey & dataRange.Cells(r, colNums(i)).Value & Chr(7) ' Use a safe delimiter
        Next i
        If Not uniqueCombinations.Exists(compositeKey) Then uniqueCombinations.Add compositeKey, 1
    Next r
    
    If uniqueCombinations.Count = 0 Then
        MsgBox "No unique combinations found to split.", vbInformation
        GoTo Cleanup
    End If

    ' --- VII. PERFORM CUSTOM SORT ON THE KEYS ---
    sortedKeys = uniqueCombinations.Keys
    CustomBubbleSort sortedKeys, yearColIndexInSplit

    ' --- VIII. CREATE SHEETS FROM THE SORTED KEYS ---
    For Each key In sortedKeys
        splitValues = Split(key, Chr(7))
        
        ' Apply filters for each part of the combination
        For i = 0 To UBound(colNums)
            dataRange.AutoFilter Field:=colNums(i), Criteria1:=splitValues(i)
        Next i
        
        ' --- IMPROVED SHEET NAMING ---
        sanitizedValue = Join(splitValues, "_") ' Join with underscore
        ' Remove trailing underscore if the last value was blank
        If Right(sanitizedValue, 1) = "_" Then
            sanitizedValue = Left(sanitizedValue, Len(sanitizedValue) - 1)
        End If
        sanitizedValue = Left(WorksheetFunction.Clean(sanitizedValue), 31) ' Clean and shorten name
        sanitizedValue = Replace(Replace(Replace(Replace(Replace(sanitizedValue, "/", "_"), "\", "_"), "*", "_"), "[", "_"), "]", "_")

        If Not SheetExists(sanitizedValue, ThisWorkbook) Then
            Set newWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            newWS.Name = sanitizedValue
            
            dataRange.SpecialCells(xlCellTypeVisible).Copy
            
            With newWS.Range("A1")
                .PasteSpecial Paste:=xlPasteColumnWidths
                .PasteSpecial Paste:=xlPasteAllUsingSourceTheme
            End With
            
            Set newWS = Nothing
        End If
        
        tempWS.AutoFilterMode = False ' Reset filter for next loop
    Next key

    MsgBox "Data has been successfully split into alphabetized sheets.", vbInformation

' --- IX. FINAL, IRONCLAD CLEANUP ---
Cleanup:
    On Error Resume Next
    Set dataRange = Nothing
    Set uniqueCombinations = Nothing
    If Not sourceWS Is Nothing Then sourceWS.Activate
    DoEvents
    Application.DisplayAlerts = False
    If Not tempWS Is Nothing Then tempWS.Delete
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

' --- HELPER SUBROUTINE FOR CUSTOM SORTING ---
Private Sub CustomBubbleSort(arr As Variant, yearColIndex As Long)
    Dim i As Long, j As Long
    Dim temp As Variant
    If UBound(arr) < LBound(arr) Then Exit Sub ' Exit if array is empty
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            ' Use a comparison function to decide if a swap is needed
            If CompareKeys(arr(j), arr(i), yearColIndex) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

' --- HELPER FUNCTION TO COMPARE TWO KEYS FOR SORTING ---
Private Function CompareKeys(key1 As Variant, key2 As Variant, yearColIndex As Long) As Boolean
    ' Returns TRUE if key1 should come BEFORE key2
    Dim v1() As String, v2() As String
    v1 = Split(key1, Chr(7))
    v2 = Split(key2, Chr(7))

    ' 1. Custom 'Year' comparison if applicable
    If yearColIndex <> -1 Then
        Dim yearVal1 As Long, yearVal2 As Long
        yearVal1 = GetYearSortValue(v1(yearColIndex))
        yearVal2 = GetYearSortValue(v2(yearColIndex))
        
        If yearVal1 <> yearVal2 Then
            CompareKeys = (yearVal1 < yearVal2)
            Exit Function
        End If
    End If
    
    ' 2. Standard alphabetical comparison for all other columns
    CompareKeys = (StrComp(Replace(key1, Chr(7), ""), Replace(key2, Chr(7), ""), vbTextCompare) < 0)
End Function

' --- HELPER FUNCTION TO ASSIGN A NUMERIC VALUE TO THE YEAR ---
Private Function GetYearSortValue(yearString As String) As Long
    Select Case UCase(Trim(yearString))
        Case "FE": GetYearSortValue = 1
        Case "SE": GetYearSortValue = 2
        Case "TE": GetYearSortValue = 3
        Case "BE": GetYearSortValue = 4
        Case Else: GetYearSortValue = 99 ' Other values go last
    End Select
End Function

' Helper function to robustly check if a worksheet exists
Private Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function


'================Split & add Sr. No=================
' Function to split the sheet based on unique combinations from one or more selected columns,
' with advanced custom sorting and improved sheet naming.
Sub SplitDataByMultipleColumns_Advanced()

    ' --- I. PRE-CHECKS & SETUP ---
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This operation can only be run on a worksheet.", vbExclamation
        Exit Sub
    End If

    Dim sourceWS As Worksheet, tempWS As Worksheet, newWS As Worksheet
    Set sourceWS = ThisWorkbook.ActiveSheet

    ' --- ADDED: FORMATTING CONSTANTS ---
    Const FONT_NAME As String = "Calibri"
    Const FONT_SIZE As Long = 11
    Const EXTRA_WIDTH_CM As Double = 0.2
    Const EXTRA_HEIGHT_CM As Double = 0.1
    
    ' --- II. OPTIMIZATION ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' --- III. VARIABLE DECLARATION ---
    Dim dataRange As Range
    Dim uniqueCombinations As Object ' Scripting.Dictionary
    Dim key As Variant, r As Long, i As Long, j As Long
    Dim columnHeaders As String, headerArray() As String, colNums() As Long
    Dim compositeKey As String, splitValues() As String, sanitizedValue As String
    Dim sortedKeys As Variant
    Dim yearColIndexInSplit As Long ' To track the position of the 'Year' column for custom sorting
    Dim lastDataRow As Long

    On Error GoTo Cleanup ' Ensure all cleanup steps are attempted

    ' --- IV. CREATE A SAFE, TEMPORARY WORKSPACE ---
    Set tempWS = ThisWorkbook.Sheets.Add(After:=sourceWS)
    tempWS.Visible = xlSheetVeryHidden
    sourceWS.Cells.Copy tempWS.Range("A1")

    ' --- V. GET USER INPUT & VALIDATE COLUMNS ---
    Set dataRange = tempWS.UsedRange
    If dataRange.Rows.Count <= 1 Then GoTo Cleanup
    
    columnHeaders = InputBox("Enter the column headers to split by, separated by commas:", _
                             "Split by Multiple Columns", "Branch, Year")
    If columnHeaders = "" Then GoTo Cleanup
    
    headerArray = Split(columnHeaders, ",")
    ReDim colNums(UBound(headerArray))
    yearColIndexInSplit = -1 ' Default to -1 (not found)
    
    For i = 0 To UBound(headerArray)
        Dim currentHeader As String
        currentHeader = Trim(headerArray(i))
        
        On Error Resume Next
        colNums(i) = Application.Match(currentHeader, dataRange.Rows(1), 0)
        On Error GoTo 0
        
        If colNums(i) = 0 Then
            MsgBox "Column header '" & currentHeader & "' not found.", vbExclamation
            GoTo Cleanup
        End If
        
        If LCase(currentHeader) = "year" Then yearColIndexInSplit = i
    Next i

    ' --- VI. GATHER UNIQUE COMBINATIONS ---
    Set uniqueCombinations = CreateObject("Scripting.Dictionary")
    
    For r = 2 To dataRange.Rows.Count
        compositeKey = ""
        For i = 0 To UBound(colNums)
            compositeKey = compositeKey & dataRange.Cells(r, colNums(i)).Value & Chr(7) ' Use a safe delimiter
        Next i
        If Not uniqueCombinations.Exists(compositeKey) Then uniqueCombinations.Add compositeKey, 1
    Next r
    
    If uniqueCombinations.Count = 0 Then
        MsgBox "No unique combinations found to split.", vbInformation
        GoTo Cleanup
    End If

    ' --- VII. PERFORM CUSTOM SORT ON THE KEYS ---
    sortedKeys = uniqueCombinations.Keys
    CustomBubbleSort sortedKeys, yearColIndexInSplit

    ' --- VIII. CREATE SHEETS FROM THE SORTED KEYS ---
    For Each key In sortedKeys
        splitValues = Split(key, Chr(7))
        
        ' Apply filters for each part of the combination
        For i = 0 To UBound(colNums)
            dataRange.AutoFilter Field:=colNums(i), Criteria1:=splitValues(i)
        Next i
        
        ' --- IMPROVED SHEET NAMING ---
        sanitizedValue = Join(splitValues, "_") ' Join with underscore
        If Right(sanitizedValue, 1) = "_" Then sanitizedValue = Left(sanitizedValue, Len(sanitizedValue) - 1)
        sanitizedValue = Left(WorksheetFunction.Clean(sanitizedValue), 31)
        sanitizedValue = Replace(Replace(Replace(Replace(Replace(sanitizedValue, "/", "_"), "\", "_"), "*", "_"), "[", "_"), "]", "_")

        If Not SheetExists(sanitizedValue, ThisWorkbook) Then
            Set newWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            newWS.Name = sanitizedValue
            
            dataRange.SpecialCells(xlCellTypeVisible).Copy
            
            With newWS.Range("A1")
                .PasteSpecial Paste:=xlPasteColumnWidths
                .PasteSpecial Paste:=xlPasteAllUsingSourceTheme
            End With
            
            ' --- ADD SR. NO. COLUMN TO THE NEW SHEET ---
            newWS.Columns("A").Insert Shift:=xlToRight
            newWS.Cells(1, 1).Value = "Sr. No"
            lastDataRow = newWS.Cells(newWS.Rows.Count, 2).End(xlUp).Row
            If lastDataRow > 1 Then
                For j = 2 To lastDataRow
                    newWS.Cells(j, 1).Value = j - 1
                Next j
            End If
            
            ' --- APPLY FULL, FINAL FORMATTING TO THE NEW SHEET ---
            ApplyStandardFormatting newWS, FONT_NAME, FONT_SIZE, EXTRA_WIDTH_CM, EXTRA_HEIGHT_CM
            
            Set newWS = Nothing
        End If
        
        tempWS.AutoFilterMode = False
    Next key

    MsgBox "Data has been successfully split and formatted into new sheets.", vbInformation

' --- IX. FINAL, IRONCLAD CLEANUP ---
Cleanup:
    On Error Resume Next
    Set dataRange = Nothing
    Set uniqueCombinations = Nothing
    If Not sourceWS Is Nothing Then sourceWS.Activate
    DoEvents
    Application.DisplayAlerts = False
    If Not tempWS Is Nothing Then tempWS.Delete
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    On Error GoTo 0
End Sub


' >>>>>>>>>> NEW HELPER SUBROUTINE FOR FORMATTING <<<<<<<<<<
Private Sub ApplyStandardFormatting(ByVal ws As Worksheet, FONT_NAME As String, FONT_SIZE As Long, EXTRA_WIDTH_CM As Double, EXTRA_HEIGHT_CM As Double)
    ' Applies a standard set of formats (font, alignment, borders, padding) to a given worksheet.
    
    ' --- Variable Declarations for this sub ---
    Dim lastRow As Long, lastCol As Long
    Dim targetRange As Range
    Dim extraWidthUnits As Double, extraHeightUnits As Double
    Dim col As Range, rw As Range

    ' Exit if the sheet is empty
    If WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Sub

    ' Automatically find the last row and column to define the data range
    lastRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Apply Core Formatting (Font, Alignment, and Borders)
    With targetRange
        .Font.Name = FONT_NAME
        .Font.Size = FONT_SIZE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Apply AutoFit
    targetRange.Columns.AutoFit
    targetRange.Rows.AutoFit

    ' Add extra padding
    If EXTRA_WIDTH_CM > 0 Then
        extraWidthUnits = EXTRA_WIDTH_CM / 0.2
        For Each col In targetRange.Columns
            col.ColumnWidth = col.ColumnWidth + extraWidthUnits
        Next col
    End If
    
    If EXTRA_HEIGHT_CM > 0 Then
        extraHeightUnits = EXTRA_HEIGHT_CM * 28.35
        For Each rw In targetRange.Rows
            rw.RowHeight = rw.RowHeight + extraHeightUnits
        Next rw
    End If
End Sub


' --- HELPER SUBROUTINE FOR CUSTOM SORTING ---
Private Sub CustomBubbleSort(arr As Variant, yearColIndex As Long)
    Dim i As Long, j As Long
    Dim temp As Variant
    If UBound(arr) < LBound(arr) Then Exit Sub
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CompareKeys(arr(j), arr(i), yearColIndex) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

' --- HELPER FUNCTION TO COMPARE TWO KEYS FOR SORTING ---
Private Function CompareKeys(key1 As Variant, key2 As Variant, yearColIndex As Long) As Boolean
    Dim v1() As String, v2() As String
    v1 = Split(key1, Chr(7))
    v2 = Split(key2, Chr(7))
    If yearColIndex <> -1 Then
        Dim yearVal1 As Long, yearVal2 As Long
        yearVal1 = GetYearSortValue(v1(yearColIndex))
        yearVal2 = GetYearSortValue(v2(yearColIndex))
        If yearVal1 <> yearVal2 Then
            CompareKeys = (yearVal1 < yearVal2)
            Exit Function
        End If
    End If
    CompareKeys = (StrComp(Replace(key1, Chr(7), ""), Replace(key2, Chr(7), ""), vbTextCompare) < 0)
End Function

' --- HELPER FUNCTION TO ASSIGN A NUMERIC VALUE TO THE YEAR ---
Private Function GetYearSortValue(yearString As String) As Long
    Select Case UCase(Trim(yearString))
        Case "FE": GetYearSortValue = 1
        Case "SE": GetYearSortValue = 2
        Case "TE": GetYearSortValue = 3
        Case "BE": GetYearSortValue = 4
        Case Else: GetYearSortValue = 99
    End Select
End Function

' Helper function to robustly check if a worksheet exists
Private Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function


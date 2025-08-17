'===============Split In Sorted Order==================
Sub CreateConsolidatedReport()

    ' --- I. SETTINGS & PRE-CHECKS ---
    Const FONT_NAME As String = "Calibri"
    Const FONT_SIZE As Long = 11
    Const EXTRA_WIDTH_CM As Double = 0.2
    Const EXTRA_HEIGHT_CM As Double = 0.1
    Const NEW_SHEET_NAME As String = "Consolidated Report"
    Const YEAR_COL As String = "Year", BRANCH_COL As String = "Branch", DIVISION_COL As String = "Division"
    Const ROLLNO_COL As String = "Roll No.", NAME_COL As String = "Name"
    Const YEAR_CUSTOM_ORDER As String = "FE,SE,TE,BE"
    
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This operation can only be run on a worksheet.", vbExclamation, "Invalid Sheet Type"
        Exit Sub
    End If

    ' --- II. VARIABLE DECLARATION ---
    Dim sourceWS As Worksheet, destWS As Worksheet
    Dim sortRange As Range, dataRange As Range
    Dim r As Long, i As Long, lastRow As Long, lastCol As Long
    Dim columnHeaders As String, headerArray() As String, colNums() As Long
    Dim currentKey As String, previousKey As String, groupName As String
    Dim serialCounter As Long
    Dim yearHeader As Range, branchHeader As Range, divHeader As Range, rollHeader As Range, nameHeader As Range
    Dim srNoChoice As VbMsgBoxResult, headerChoice As VbMsgBoxResult

    ' --- III. OPTIMIZATION ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo Cleanup

    ' --- IV. SETUP SOURCE AND DESTINATION SHEETS ---
    Set sourceWS = ThisWorkbook.ActiveSheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(NEW_SHEET_NAME).Delete
    On Error GoTo Cleanup
    Application.DisplayAlerts = True
    Set destWS = ThisWorkbook.Sheets.Add(After:=sourceWS)
    destWS.Name = NEW_SHEET_NAME
    sourceWS.UsedRange.Copy destWS.Range("A1")
    
    ' --- V. PERFORM FIXED, MULTI-LEVEL SORT ---
    lastRow = destWS.Cells(destWS.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 1 Then GoTo Cleanup
    lastCol = destWS.Cells(1, destWS.Columns.Count).End(xlToLeft).Column
    Set sortRange = destWS.Range("A1").Resize(lastRow, lastCol)
    On Error Resume Next
    Set yearHeader = destWS.Rows(1).Find(What:=YEAR_COL, LookAt:=xlWhole, MatchCase:=False)
    Set branchHeader = destWS.Rows(1).Find(What:=BRANCH_COL, LookAt:=xlWhole, MatchCase:=False)
    Set divHeader = destWS.Rows(1).Find(What:=DIVISION_COL, LookAt:=xlWhole, MatchCase:=False)
    Set rollHeader = destWS.Rows(1).Find(What:=ROLLNO_COL, LookAt:=xlWhole, MatchCase:=False)
    Set nameHeader = destWS.Rows(1).Find(What:=NAME_COL, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo Cleanup
    If yearHeader Is Nothing Or branchHeader Is Nothing Or divHeader Is Nothing Or rollHeader Is Nothing Or nameHeader Is Nothing Then
        MsgBox "One or more required columns for sorting could not be found.", vbCritical, "Columns Not Found"
        GoTo Cleanup
    End If
    With destWS.Sort
        .SortFields.Clear
        .SortFields.Add Key:=yearHeader, SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=YEAR_CUSTOM_ORDER
        .SortFields.Add Key:=branchHeader, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=divHeader, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=rollHeader, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=nameHeader, SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange sortRange
        .Header = xlYes: .MatchCase = False: .Orientation = xlTopToBottom
        .Apply
    End With

    ' --- VI. GET USER INPUT FOR GROUPING SEPARATORS ---
    columnHeaders = InputBox("Data is now sorted." & vbCrLf & vbCrLf & _
                             "Enter column headers to define a 'group' (e.g., Branch, Year):", _
                             "Define Group Separators", "Branch, Year")
    If columnHeaders = "" Then GoTo Cleanup
    headerArray = Split(columnHeaders, ",")
    ReDim colNums(UBound(headerArray))
    Set dataRange = destWS.UsedRange
    For i = 0 To UBound(headerArray)
        Dim foundHeader As Range
        Set foundHeader = dataRange.Rows(1).Find(What:=Trim(headerArray(i)), LookAt:=xlWhole, MatchCase:=False)
        If foundHeader Is Nothing Then
            MsgBox "Grouping column '" & Trim(headerArray(i)) & "' not found. Aborting.", vbCritical
            GoTo Cleanup
        Else
            colNums(i) = foundHeader.Column
        End If
    Next i
    
    ' --- VII. ASK USER PREFERENCES (GROUP HEADERS & SR. NO.) ---
    headerChoice = MsgBox("Do you want to add titled, colored headers for each group?", _
                          vbYesNo + vbQuestion, "Add Group Headers?")
                          
    srNoChoice = MsgBox("Do you want to add Serial Numbers?" & vbCrLf & vbCrLf & _
                        "• YES = Restart numbering for each group" & vbCrLf & _
                        "• NO = Use continuous numbering (1, 2, 3...)" & vbCrLf & _
                        "• CANCEL = Do not add a Sr. No. column", _
                        vbYesNoCancel + vbQuestion, "Add Serial Numbers?")
    
    If srNoChoice <> vbCancel Then
        destWS.Columns("A").Insert Shift:=xlToRight
        destWS.Cells(1, 1).Value = "Sr. No"
        lastCol = destWS.Cells(1, destWS.Columns.Count).End(xlToLeft).Column
    End If
    
    ' --- VIII. INSERT ROWS BASED ON USER CHOICES ---
    lastRow = destWS.Cells(destWS.Rows.Count, IIf(srNoChoice <> vbCancel, 2, 1)).End(xlUp).Row
    For r = lastRow To 3 Step -1
        currentKey = "": previousKey = ""
        Dim offset As Integer: offset = IIf(srNoChoice <> vbCancel, 1, 0)
        For i = 0 To UBound(colNums)
            currentKey = currentKey & destWS.Cells(r, colNums(i) + offset).Value & Chr(7)
            previousKey = previousKey & destWS.Cells(r - 1, colNums(i) + offset).Value & Chr(7)
        Next i
        
        If currentKey <> previousKey Then
            If headerChoice = vbYes Then
                ' --- INSERT SEPARATOR AND A FORMATTED HEADER ---
                destWS.Rows(r).Insert Shift:=xlDown
                destWS.Rows(r).Insert Shift:=xlDown
                groupName = Replace(currentKey, Chr(7), "-")
                If Right(groupName, 1) = "-" Then groupName = Left(groupName, Len(groupName) - 1)
                With destWS.Range(destWS.Cells(r + 1, 1), destWS.Cells(r + 1, lastCol))
                    .Merge
                    .Value = UCase(groupName)
                    .Interior.Color = RGB(198, 239, 206) ' Light Green
                    .Font.Name = "Calibri"
                    .Font.Size = 11
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            Else
                ' --- INSERT A SIMPLE BLANK SEPARATOR ROW ---
                destWS.Rows(r).Insert Shift:=xlDown
            End If
        End If
    Next r

    ' --- IX. POPULATE "SR. NO." COLUMN (IF CHOSEN) ---
    If srNoChoice <> vbCancel Then
        lastRow = destWS.Cells(destWS.Rows.Count, 2).End(xlUp).Row
        serialCounter = 1
        If srNoChoice = vbNo Then ' Continuous Numbering
            For r = 2 To lastRow
                If Not IsEmpty(destWS.Cells(r, 2).Value) And Not destWS.Cells(r, 1).MergeCells Then
                    destWS.Cells(r, 1).Value = serialCounter
                    serialCounter = serialCounter + 1
                End If
            Next r
        Else ' VbYes - Restarting Numbering
            previousKey = "###START###"
            For r = 2 To lastRow
                If Not IsEmpty(destWS.Cells(r, 2).Value) And Not destWS.Cells(r, 1).MergeCells Then
                    currentKey = ""
                    For i = 0 To UBound(colNums)
                        currentKey = currentKey & destWS.Cells(r, colNums(i) + 1).Value
                    Next i
                    If currentKey <> previousKey Then
                        serialCounter = 1
                        previousKey = currentKey
                    End If
                    destWS.Cells(r, 1).Value = serialCounter
                    serialCounter = serialCounter + 1
                End If
            Next r
        End If
    End If
    
    ' --- X. APPLY FINAL FORMATTING ---
    ApplyStandardFormattingToSheet destWS, lastCol, FONT_NAME, FONT_SIZE, EXTRA_WIDTH_CM, EXTRA_HEIGHT_CM

    ' --- XI. FINALIZATION ---
    destWS.Activate
    destWS.Cells(1, 1).Select
    MsgBox "Consolidated report has been created successfully!", vbInformation, "Process Complete"

' --- CLEANUP BLOCK ---
Cleanup:
    If Err.Number <> 0 Then MsgBox "An error occurred: " & vbCrLf & Err.Description, vbCritical, "Macro Error"
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True: Application.DisplayAlerts = True
End Sub


' --- HELPER SUBROUTINE FOR APPLYING FORMAT ---
Private Sub ApplyStandardFormattingToSheet(ByVal ws As Worksheet, ByVal lastCol As Long, FONT_NAME As String, FONT_SIZE As Long, EXTRA_WIDTH_CM As Double, EXTRA_HEIGHT_CM As Double)
    Dim targetRange As Range, extraWidthUnits As Double, extraHeightUnits As Double, col As Range, rw As Range
    If WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Sub
    Set targetRange = ws.UsedRange
    targetRange.Columns.AutoFit
    If EXTRA_WIDTH_CM > 0 Then
        extraWidthUnits = EXTRA_WIDTH_CM / 0.2
        For Each col In ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Columns
            col.ColumnWidth = col.ColumnWidth + extraWidthUnits
        Next col
    End If
    If EXTRA_HEIGHT_CM > 0 Then extraHeightUnits = EXTRA_HEIGHT_CM * 28.35
    For Each rw In targetRange.Rows
        If WorksheetFunction.CountA(rw) > 0 Then
            ' Group Header rows are already formatted, just ensure height.
            If rw.Cells(1, 1).MergeCells Then
                rw.RowHeight = 22
            Else 'This is a data row
                With rw.Resize(, lastCol)
                    .Font.Name = FONT_NAME: .Font.Size = FONT_SIZE
                    .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                    .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
                    .AutoFit
                    If EXTRA_HEIGHT_CM > 0 Then .RowHeight = .RowHeight + extraHeightUnits
                End With
            End If
        Else ' This is a blank separator row
            With rw
                .Borders.LineStyle = xlNone
                .RowHeight = 20
            End With
        End If
    Next rw
End Sub




'===============Split In Non-Sorted Order==================================
Sub SplitDataWithRowGaps()

    ' --- I. SETTINGS & PRE-CHECKS ---
    ' Formatting Constants
    Const FONT_NAME As String = "Calibri"
    Const FONT_SIZE As Long = 11
    Const EXTRA_WIDTH_CM As Double = 0.2
    Const EXTRA_HEIGHT_CM As Double = 0.1
    
    ' Sheet Naming
    Const NEW_SHEET_NAME As String = "Consolidated Report"

    ' Check if we are on a valid worksheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This operation can only be run on a worksheet.", vbExclamation, "Invalid Sheet Type"
        Exit Sub
    End If

    ' --- II. VARIABLE DECLARATION ---
    Dim sourceWS As Worksheet, destWS As Worksheet
    Dim dataRange As Range, sortRange As Range
    Dim r As Long, i As Long
    Dim lastRow As Long
    Dim columnHeaders As String, headerArray() As String, colNums() As Long
    Dim currentKey As String, previousKey As String
    Dim serialCounter As Long

    ' --- III. OPTIMIZATION ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo Cleanup ' Ensure settings are restored even if an error occurs

    ' --- IV. SETUP SOURCE AND DESTINATION SHEETS ---
    Set sourceWS = ThisWorkbook.ActiveSheet

    ' Check if a report sheet already exists and delete it to start fresh
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(NEW_SHEET_NAME).Delete
    On Error GoTo Cleanup
    Application.DisplayAlerts = True
    
    ' Create the new destination sheet and copy data
    Set destWS = ThisWorkbook.Sheets.Add(After:=sourceWS)
    destWS.Name = NEW_SHEET_NAME
    sourceWS.UsedRange.Copy destWS.Range("A1")
    
    ' --- V. GET USER INPUT & VALIDATE COLUMNS ---
    Set dataRange = destWS.UsedRange
    If dataRange.Rows.Count <= 1 Then
        MsgBox "No data to process.", vbInformation
        GoTo Cleanup
    End If
    
    columnHeaders = InputBox("Enter the column headers to group by, separated by commas:", _
                             "Group and Separate Data", "Branch, Year")
    If columnHeaders = "" Then GoTo Cleanup ' User cancelled
    
    headerArray = Split(columnHeaders, ",")
    ReDim colNums(UBound(headerArray))
    
    ' Find the column number for each header
    For i = 0 To UBound(headerArray)
        Dim foundHeader As Range
        Set foundHeader = dataRange.Rows(1).Find(What:=Trim(headerArray(i)), LookAt:=xlWhole, MatchCase:=False)
        
        If foundHeader Is Nothing Then
            MsgBox "Column header '" & Trim(headerArray(i)) & "' not found. Aborting.", vbCritical
            GoTo Cleanup
        Else
            colNums(i) = foundHeader.Column
        End If
    Next i

    ' --- VI. SORT DATA ON THE NEW SHEET ---
    lastRow = destWS.Cells(destWS.Rows.Count, 1).End(xlUp).Row
    Set sortRange = destWS.Range("A1").Resize(lastRow, destWS.UsedRange.Columns.Count)

    With destWS.Sort
        .SortFields.Clear
        For i = 0 To UBound(colNums)
            If LCase(Trim(headerArray(i))) = "year" Then
                .SortFields.Add Key:=destWS.Columns(colNums(i)), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="FE,SE,TE,BE", DataOption:=xlSortNormal
            Else
                .SortFields.Add Key:=destWS.Columns(colNums(i)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            End If
        Next i
        
        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' --- VII. INSERT BLANK ROWS BETWEEN GROUPS ---
    lastRow = destWS.Cells(destWS.Rows.Count, 1).End(xlUp).Row
    For r = lastRow To 3 Step -1
        currentKey = ""
        previousKey = ""
        For i = 0 To UBound(colNums)
            currentKey = currentKey & destWS.Cells(r, colNums(i)).Value
            previousKey = previousKey & destWS.Cells(r - 1, colNums(i)).Value
        Next i
        
        If currentKey <> previousKey Then
            destWS.Rows(r).Insert Shift:=xlDown
        End If
    Next r

    ' --- VIII. ADD AND POPULATE "SR. NO." COLUMN ---
    destWS.Columns("A").Insert Shift:=xlToRight
    destWS.Cells(1, 1).Value = "Sr. No"
    serialCounter = 1
    
    lastRow = destWS.Cells(destWS.Rows.Count, 2).End(xlUp).Row
    
    For r = 2 To lastRow
        If Not IsEmpty(destWS.Cells(r, 2).Value) Then
            destWS.Cells(r, 1).Value = serialCounter
            serialCounter = serialCounter + 1
        End If
    Next r

    ' --- IX. APPLY FINAL FORMATTING ---
    ApplyStandardFormattingToSheet destWS, FONT_NAME, FONT_SIZE, EXTRA_WIDTH_CM, EXTRA_HEIGHT_CM

    ' --- X. FINALIZATION ---
    destWS.Activate
    destWS.Cells(1, 1).Select
    MsgBox "Consolidated report has been created successfully!", vbInformation, "Process Complete"

' --- CLEANUP BLOCK ---
Cleanup:
    If Err.Number <> 0 Then
        MsgBox "An error occurred: " & vbCrLf & Err.Description, vbCritical, "Macro Error"
    End If
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


' --- HELPER SUBROUTINE FOR APPLYING FORMAT (REVISED LOGIC) ---
Private Sub ApplyStandardFormattingToSheet(ByVal ws As Worksheet, FONT_NAME As String, FONT_SIZE As Long, EXTRA_WIDTH_CM As Double, EXTRA_HEIGHT_CM As Double)
    ' Applies specific formats to data rows while setting a clean, unformatted style for blank separator rows.
    
    Dim targetRange As Range
    Dim extraWidthUnits As Double, extraHeightUnits As Double
    Dim col As Range, rw As Range
    
    If WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Sub

    Set targetRange = ws.UsedRange
    
    ' First, apply column-wide formatting (AutoFit and extra width)
    targetRange.Columns.AutoFit
    If EXTRA_WIDTH_CM > 0 Then
        extraWidthUnits = EXTRA_WIDTH_CM / 0.2
        For Each col In targetRange.Columns
            col.ColumnWidth = col.ColumnWidth + extraWidthUnits
        Next col
    End If

    ' Now, loop row-by-row to apply different formatting to data rows vs. blank separator rows.
    If EXTRA_HEIGHT_CM > 0 Then
        extraHeightUnits = EXTRA_HEIGHT_CM * 28.35
    End If
    
    For Each rw In targetRange.Rows
        ' Check if the row contains any data. CountA is perfect for this.
        If WorksheetFunction.CountA(rw) > 0 Then
            ' --- THIS IS A DATA ROW ---
            With rw
                ' Apply Core Formatting (Font, Alignment, and Borders)
                .Font.Name = FONT_NAME
                .Font.Size = FONT_SIZE
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                
                ' Apply AutoFit and extra height
                .AutoFit
                If EXTRA_HEIGHT_CM > 0 Then
                    .RowHeight = .RowHeight + extraHeightUnits
                End If
            End With
        Else
            ' --- THIS IS A BLANK SEPARATOR ROW ---
            With rw
                ' Remove any borders that might have been accidentally applied
                .Borders.LineStyle = xlNone
                ' Set a fixed, small height for a clean visual gap
                .RowHeight = 25
            End With
        End If
    Next rw
End Sub
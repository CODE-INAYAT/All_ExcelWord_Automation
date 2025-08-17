Sub ApplyFullFormatting()

    ' --- SETTINGS ---
    ' This macro will run on the currently ACTIVE sheet.
    
    ' Font and Size Settings
    Const FONT_NAME As String = "Calibri"
    Const FONT_SIZE As Long = 11
    
    ' Padding Settings (in cm)
    Const EXTRA_WIDTH_CM As Double = 0.2
    Const EXTRA_HEIGHT_CM As Double = 0.1
    ' ----------------
    
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim extraWidthUnits As Double
    Dim extraHeightUnits As Double
    Dim col As Range
    Dim rw As Range ' Using "rw" to avoid conflict with the "Row" keyword

    ' 1. Validate and set the worksheet object to the active sheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This macro must be run on a standard worksheet.", vbExclamation, "Invalid Sheet Type"
        Exit Sub
    End If
    Set ws = ActiveSheet

    ' Turn off screen updating for a massive performance boost
    Application.ScreenUpdating = False

    ' 2. Automatically find the last row and column to define the data range
    On Error Resume Next
    lastRow = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0
    
    If lastRow = 0 Or lastCol = 0 Then
        MsgBox "No data found on the active sheet to format.", vbInformation, "Sheet is Empty"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' 3. Apply Core Formatting (Font, Alignment, and Borders)
    With targetRange
        ' *** NEW: Set Font and Font Size for the entire range ***
        .Font.Name = FONT_NAME
        .Font.Size = FONT_SIZE
        
        ' Set Alignment
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' Apply "All Borders"
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' 4. Apply AutoFit to the entire detected range
    targetRange.Columns.AutoFit
    targetRange.Rows.AutoFit

    ' 5. [Optional] Add extra padding for better readability
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
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True

    ' Update the confirmation message
    MsgBox "Full formatting has been applied (Font, Alignment, Borders, AutoFit).", vbInformation, "Formatting Complete"
End Sub


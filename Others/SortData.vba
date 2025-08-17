Sub SortStudentData()

    Const YEAR_COL As String = "Year"
    Const BRANCH_COL As String = "Branch"
    Const DIVISION_COL As String = "Division"
    Const ROLLNO_COL As String = "Roll No."
    Const NAME_COL As String = "Name"
    
    Const YEAR_CUSTOM_ORDER As String = "FE,SE,TE,BE"

    Dim ws As Worksheet
    Dim sortRange As Range
    Dim lastRow As Long
    Dim lastCol As Long

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "This macro must be run on a standard worksheet.", vbExclamation, "Invalid Sheet Type"
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then
        MsgBox "No data found to sort on the active sheet.", vbInformation, "No Data"
        Exit Sub
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set sortRange = ws.Range("A1").Resize(lastRow, lastCol)

    On Error Resume Next
    Dim yearHeader As Range, branchHeader As Range, divHeader As Range
    Dim rollHeader As Range, nameHeader As Range
    
    Set yearHeader = ws.Rows(1).Find(What:=YEAR_COL, LookAt:=xlWhole, MatchCase:=False)
    Set branchHeader = ws.Rows(1).Find(What:=BRANCH_COL, LookAt:=xlWhole, MatchCase:=False)
    Set divHeader = ws.Rows(1).Find(What:=DIVISION_COL, LookAt:=xlWhole, MatchCase:=False)
    Set rollHeader = ws.Rows(1).Find(What:=ROLLNO_COL, LookAt:=xlWhole, MatchCase:=False)
    Set nameHeader = ws.Rows(1).Find(What:=NAME_COL, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If yearHeader Is Nothing Or branchHeader Is Nothing Or divHeader Is Nothing Or rollHeader Is Nothing Or nameHeader Is Nothing Then
        MsgBox "One or more required columns could not be found. Please ensure the active sheet has headers named: " & _
               vbCrLf & "- " & YEAR_COL & vbCrLf & "- " & BRANCH_COL & vbCrLf & "- " & DIVISION_COL & _
               vbCrLf & "- " & ROLLNO_COL & vbCrLf & "- " & NAME_COL, vbCritical, "Columns Not Found"
        Exit Sub
    End If

    With ws.Sort
        .SortFields.Clear
        
        .SortFields.Add Key:=yearHeader, SortOn:=xlSortOnValues, Order:=xlAscending, _
                        CustomOrder:=YEAR_CUSTOM_ORDER, DataOption:=xlSortNormal
                        
        .SortFields.Add Key:=branchHeader, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
        .SortFields.Add Key:=divHeader, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
        .SortFields.Add Key:=rollHeader, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
        .SortFields.Add Key:=nameHeader, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Application.ScreenUpdating = True
    
    MsgBox "Data has been successfully sorted!", vbInformation, "Task Complete"

End Sub
' This code marks specific T&P UIDs as present and highlights rows where "P" is there
Sub MarkAndHighlightPresent()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim attendanceCol As Long, uidCol As Long
    Dim i As Long, j As Long
    Dim uidColumnName As String
    Dim coreFoundUIDs As String, coreNotFoundUIDs As String
    Dim wcFoundUIDs As String, wcNotFoundUIDs As String
    Dim coreUIDList As Variant, coreNameList As Variant
    Dim wcUIDList As Variant, wcNameList As Variant
    Dim found As Boolean
    
    ' Set worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Define CORE T&P UIDs and Names
    coreUIDList = Array("23-E&CSA59-27", "24-COMPB66-27", "24-CS&E57-28", "23-COMPA12-27", _
                        "23-AI&DSB24-27", "23-CSEA55-27", "23-E&CSA58-27", "24-ITC67-27", _
                        "23-COMPA20-27", "23-COMPA54-27")
    
    coreNameList = Array("Devansh", "Prince", "Rudra", "Hardik", _
                         "Inayatulla", "Sneha", "Ayush", "Ansh", _
                         "Nidhay Chavan", "Karan")
    
    ' Define WC T&P UIDs and Names
    wcUIDList = Array("24-AI&MLB24-28", "24-AI&MLB20-28", "24-ITA51-28", "24-CS&E20-28", _
                      "24-COMPA54-28", "24-AI&DSB28-28", "24-COMPA44-28", "24-COMPA17-28", _
                      "24-CSEA65-27", "24-AI&MLB14-28")
    
    wcNameList = Array("Nikita Mishra", "Varun Maurya", "Ravishankar Kanaki", "Tanishka Jaiswal", _
                       "Ajitesh Jain", "Adhithya Nair", "Himanshu Gupta", "Sanju Chauhan", _
                       "Tavleen Kaur Dadial", "Alok Sharad Mahadik")
    
    ' Ask user for the T&P UID column name
    uidColumnName = InputBox("Enter the column name for T&P UID (default: T&P UID):", _
                            "Column Name Input", "T&P UID")
    
    If uidColumnName = "" Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Find the "Attendance" column number
    On Error Resume Next
    attendanceCol = Application.WorksheetFunction.Match("Attendance", ws.Rows(1), 0)
    On Error GoTo 0
    
    If attendanceCol = 0 Then
        MsgBox "Attendance column not found in Row 1.", vbExclamation
        Exit Sub
    End If
    
    ' Find the T&P UID column number
    On Error Resume Next
    uidCol = Application.WorksheetFunction.Match(uidColumnName, ws.Rows(1), 0)
    On Error GoTo 0
    
    If uidCol = 0 Then
        MsgBox "Column '" & uidColumnName & "' not found in Row 1.", vbExclamation
        Exit Sub
    End If
    
    ' Find the last used row and column
    lastRow = ws.Cells(ws.Rows.Count, attendanceCol).End(xlUp).Row
    If lastRow < 2 Then lastRow = ws.Cells(ws.Rows.Count, uidCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Initialize report strings
    coreFoundUIDs = ""
    coreNotFoundUIDs = ""
    wcFoundUIDs = ""
    wcNotFoundUIDs = ""
    
    ' Process CORE UIDs
    For j = 0 To UBound(coreUIDList)
        found = False
        
        ' Search for the UID in the worksheet
        For i = 2 To lastRow
            If Trim(ws.Cells(i, uidCol).Value) = coreUIDList(j) Then
                ' Mark as present
                ws.Cells(i, attendanceCol).Value = "P"
                coreFoundUIDs = coreFoundUIDs & "  • " & coreNameList(j) & " (" & coreUIDList(j) & ")" & vbCrLf
                found = True
                Exit For
            End If
        Next i
        
        ' If not found, add to not found list
        If Not found Then
            coreNotFoundUIDs = coreNotFoundUIDs & "  • " & coreNameList(j) & " (" & coreUIDList(j) & ")" & vbCrLf
        End If
    Next j
    
    ' Process WC UIDs
    For j = 0 To UBound(wcUIDList)
        found = False
        
        ' Search for the UID in the worksheet
        For i = 2 To lastRow
            If Trim(ws.Cells(i, uidCol).Value) = wcUIDList(j) Then
                ' Mark as present
                ws.Cells(i, attendanceCol).Value = "P"
                wcFoundUIDs = wcFoundUIDs & "  • " & wcNameList(j) & " (" & wcUIDList(j) & ")" & vbCrLf
                found = True
                Exit For
            End If
        Next i
        
        ' If not found, add to not found list
        If Not found Then
            wcNotFoundUIDs = wcNotFoundUIDs & "  • " & wcNameList(j) & " (" & wcUIDList(j) & ")" & vbCrLf
        End If
    Next j
    
    ' Highlight rows with "P" in Attendance column
    For i = 2 To lastRow
        If Trim(ws.Cells(i, attendanceCol).Value) = "P" Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Interior.Color = RGB(198, 239, 206) ' Light green
        End If
    Next i
    
    ' Build the report with sections
    Dim reportMsg As String
    reportMsg = "ATTENDANCE MARKING REPORT" & vbCrLf & _
                "=========================" & vbCrLf & vbCrLf
    
    reportMsg = reportMsg & "CORE Names Present:" & vbCrLf
    If coreFoundUIDs = "" Then
        reportMsg = reportMsg & "  None" & vbCrLf
    Else
        reportMsg = reportMsg & coreFoundUIDs
    End If
    
    reportMsg = reportMsg & vbCrLf & "CORE Names (Skipped):" & vbCrLf
    If coreNotFoundUIDs = "" Then
        reportMsg = reportMsg & "  None" & vbCrLf
    Else
        reportMsg = reportMsg & coreNotFoundUIDs
    End If
    
    reportMsg = reportMsg & vbCrLf & "WC Names Present:" & vbCrLf
    If wcFoundUIDs = "" Then
        reportMsg = reportMsg & "  None" & vbCrLf
    Else
        reportMsg = reportMsg & wcFoundUIDs
    End If
    
    reportMsg = reportMsg & vbCrLf & "WC Names (Skipped):" & vbCrLf
    If wcNotFoundUIDs = "" Then
        reportMsg = reportMsg & "  None" & vbCrLf
    Else
        reportMsg = reportMsg & wcNotFoundUIDs
    End If
    
    ' Create a temporary text file for the report
    Dim fso As Object, txtFile As Object
    Dim tempPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = Environ("TEMP") & "\AttendanceReport.txt"
    Set txtFile = fso.CreateTextFile(tempPath, True)
    txtFile.WriteLine reportMsg
    txtFile.Close
    
    ' Show summary and ask if user wants to see detailed report
    MsgBox "Attendance marking completed!" & vbCrLf & vbCrLf & _
           "Rows with 'P' in Attendance column are highlighted in light green." & vbCrLf & vbCrLf & _
           "A detailed report has been saved to: " & tempPath & vbCrLf & vbCrLf & _
           "You can open this file to see which UIDs were marked present and which were not found.", vbInformation
    
    ' Optional: Open the report file automatically
    If MsgBox("Do you want to open the detailed report now?", vbYesNo + vbQuestion) = vbYes Then
        Shell "notepad.exe " & tempPath, vbNormalFocus
    End If
End Sub
'===================Multi Find & Replace======================'
Option Explicit

' --- The Main Macro to Launch the Tool ---
Public Sub ShowMultiReplaceForm()
    Const FORM_NAME As String = "frmMultiReplace"
    
    Dim VBE As Object         ' Late binding for VBIDE.VBE
    Dim vbProj As Object      ' Late binding for VBIDE.VBProject
    Dim comp As Object        ' Late binding for VBIDE.VBComponent
    Dim formExists As Boolean

    ' Use late binding to avoid needing a manual reference to the "Extensibility" library.
    On Error GoTo VbeError
    Set VBE = Application.VBE
    On Error GoTo 0
    Set vbProj = VBE.ActiveVBProject

    ' Check if the UserForm component already exists in the project.
    formExists = False
    On Error Resume Next ' In case project is protected and components can't be read
    For Each comp In vbProj.VBComponents
        If comp.Name = FORM_NAME Then
            formExists = True
            Exit For
        End If
    Next comp
    On Error GoTo 0

    ' If the form doesn't exist, create it.
    If Not formExists Then
        If Not CreateMultiReplaceForm(FORM_NAME, vbProj) Then
            ' If creation failed, the Create function will have already shown a message.
            Exit Sub
        End If
    End If
    
    ' Show the form modelessly, which allows interaction with the Excel sheet.
    VBA.UserForms.Add(FORM_NAME).Show vbModeless
    Exit Sub
    
VbeError:
    MsgBox "Could not access the VBA project." & vbCrLf & vbCrLf & _
           "Please ensure 'Trust access to the VBA project object model' is enabled in:" & vbCrLf & _
           "File > Options > Trust Center > Trust Center Settings > Macro Settings.", _
           vbCritical, "Access Denied"
End Sub


' --- The Function to Dynamically Create the UserForm ---
Private Function CreateMultiReplaceForm(ByVal formName As String, ByVal vbProj As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim uf As Object ' VBComponent
    Dim ctrl As Object ' MSForms.Control
    Dim codeModule As Object ' CodeModule

    ' === PART 1: Create the UserForm Object and Set Properties ===
    Set uf = vbProj.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    With uf
        .Name = formName
        .Properties("Caption") = "Multi Find & Replace"
        .Properties("Width") = 340
        .Properties("Height") = 420
    End With

    ' === PART 2: Add Controls to the UserForm ===
    ' --- Labels ---
    Set ctrl = uf.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Top = 12: .Left = 12: .Width = 60: .Caption = "Find what:": End With
    Set ctrl = uf.Designer.Controls.Add("Forms.Label.1")
    With ctrl: .Top = 36: .Left = 12: .Width = 60: .Caption = "Replace with:": End With
    
    ' --- TextBoxes for Find and Replace ---
    Set ctrl = uf.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtFind": .Top = 10: .Left = 80: .Width = 150: .Height = 18: End With
    Set ctrl = uf.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl: .Name = "txtReplace": .Top = 34: .Left = 80: .Width = 150: .Height = 18: End With

    ' --- Buttons for List Management ---
    Set ctrl = uf.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnAdd": .Top = 10: .Left = 240: .Width = 80: .Height = 22: .Caption = "Add to List": End With
    Set ctrl = uf.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnRemove": .Top = 140: .Left = 240: .Width = 80: .Height = 22: .Caption = "Remove Selected": End With
    Set ctrl = uf.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClear": .Top = 165: .Left = 240: .Width = 80: .Height = 22: .Caption = "Clear List": End With

    ' --- ListBox to display pairs ---
    Set ctrl = uf.Designer.Controls.Add("Forms.ListBox.1")
    With ctrl
        .Name = "lstPairs"
        .Top = 60: .Left = 12: .Width = 218: .Height = 130
        .ColumnCount = 2
        .ColumnWidths = "100;100"
    End With

    ' --- Options Frame and CheckBoxes ---
    Set ctrl = uf.Designer.Controls.Add("Forms.Frame.1")
    With ctrl: .Name = "fraOptions": .Caption = " Options ": .Top = 200: .Left = 12: .Width = 310: .Height = 55: End With
    Set ctrl = uf.Designer.Controls.Add("Forms.CheckBox.1")
    With ctrl: .Name = "chkMatchCase": .Caption = "Match case": .Top = 220: .Left = 24: .Width = 100: End With
    Set ctrl = uf.Designer.Controls.Add("Forms.CheckBox.1")
    With ctrl: .Name = "chkMatchEntire": .Caption = "Match entire cell contents": .Top = 220: .Left = 150: .Width = 160: End With

    ' --- Scope Frame and OptionButtons ---
    Set ctrl = uf.Designer.Controls.Add("Forms.Frame.1")
    With ctrl: .Name = "fraScope": .Caption = " Scope ": .Top = 260: .Left = 12: .Width = 310: .Height = 55: End With
    Set ctrl = uf.Designer.Controls.Add("Forms.OptionButton.1")
    With ctrl: .Name = "optSelection": .Caption = "Selection": .Top = 280: .Left = 24: .Width = 70: End With
    Set ctrl = uf.Designer.Controls.Add("Forms.OptionButton.1")
    With ctrl: .Name = "optActiveSheet": .Caption = "Active Sheet": .Top = 280: .Left = 100: .Width = 90: .Value = True: End With
    Set ctrl = uf.Designer.Controls.Add("Forms.OptionButton.1")
    With ctrl: .Name = "optWorkbook": .Caption = "Entire Workbook": .Top = 280: .Left = 200: .Width = 110: End With

    ' --- Main Action Buttons ---
    Set ctrl = uf.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnReplaceAll": .Caption = "Replace All": .Top = 325: .Left = 160: .Width = 80: .Height = 25: .Default = True: End With
    Set ctrl = uf.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl: .Name = "btnClose": .Caption = "Close": .Top = 325: .Left = 245: .Width = 80: .Height = 25: .Cancel = True: End With

    ' === PART 3: Inject the VBA Code into the UserForm's Code Module ===
    Set codeModule = uf.CodeModule
    Dim sCode As String
    
    sCode = "Private Sub btnAdd_Click()" & vbCrLf
    sCode = sCode & "    If Trim(Me.txtFind.Value) = vbNullString Then" & vbCrLf
    sCode = sCode & "        MsgBox ""'Find what' text cannot be empty."", vbExclamation, ""Input Required""" & vbCrLf
    sCode = sCode & "        Me.txtFind.SetFocus" & vbCrLf
    sCode = sCode & "        Exit Sub" & vbCrLf
    sCode = sCode & "    End If" & vbCrLf
    sCode = sCode & "    With Me.lstPairs" & vbCrLf
    sCode = sCode & "        .AddItem" & vbCrLf
    sCode = sCode & "        .List(.ListCount - 1, 0) = Me.txtFind.Value" & vbCrLf
    sCode = sCode & "        .List(.ListCount - 1, 1) = Me.txtReplace.Value" & vbCrLf
    sCode = sCode & "    End With" & vbCrLf
    sCode = sCode & "    Me.txtFind.Value = vbNullString" & vbCrLf
    sCode = sCode & "    Me.txtReplace.Value = vbNullString" & vbCrLf
    sCode = sCode & "    Me.txtFind.SetFocus" & vbCrLf
    sCode = sCode & "End Sub" & vbCrLf & vbCrLf
    
    sCode = sCode & "Private Sub btnRemove_Click()" & vbCrLf
    sCode = sCode & "    If Me.lstPairs.ListIndex > -1 Then" & vbCrLf
    sCode = sCode & "        Me.lstPairs.RemoveItem Me.lstPairs.ListIndex" & vbCrLf
    sCode = sCode & "    Else" & vbCrLf
    sCode = sCode & "        MsgBox ""Please select an item from the list to remove."", vbInformation, ""No Item Selected""" & vbCrLf
    sCode = sCode & "    End If" & vbCrLf
    sCode = sCode & "End Sub" & vbCrLf & vbCrLf

    sCode = sCode & "Private Sub btnClear_Click()" & vbCrLf
    sCode = sCode & "    Me.lstPairs.Clear" & vbCrLf
    sCode = sCode & "End Sub" & vbCrLf & vbCrLf

    sCode = sCode & "Private Sub btnClose_Click()" & vbCrLf
    sCode = sCode & "    Unload Me" & vbCrLf
    sCode = sCode & "End Sub" & vbCrLf & vbCrLf

    sCode = sCode & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    sCode = sCode & "    If CloseMode = vbFormControlMenu Then" & vbCrLf
    sCode = sCode & "        Cancel = True" & vbCrLf
    sCode = sCode & "        Unload Me" & vbCrLf
    sCode = sCode & "    End If" & vbCrLf
    sCode = sCode & "End Sub" & vbCrLf & vbCrLf

    sCode = sCode & "Private Sub btnReplaceAll_Click()" & vbCrLf
    ' *** NEW SIMPLIFIED LOGIC BLOCK START ***
    sCode = sCode & "    ' --- Check for un-added items in the text boxes before proceeding ---" & vbCrLf
    sCode = sCode & "    If Trim(Me.txtFind.Value) <> vbNullString Then" & vbCrLf
    sCode = sCode & "        Dim userResponse As VbMsgBoxResult" & vbCrLf
    sCode = sCode & "        userResponse = MsgBox(""The input fields contain text that has not been added to the list."" & vbCrLf & vbCrLf & _" & vbCrLf
    sCode = sCode & "                              ""Do you want to add it now?"", _" & vbCrLf
    sCode = sCode & "                              vbYesNo + vbQuestion, ""Un-added Item"")" & vbCrLf & vbCrLf
    sCode = sCode & "        If userResponse = vbYes Then" & vbCrLf
    sCode = sCode & "            ' Call the existing Add button's logic to add the item and clear fields" & vbCrLf
    sCode = sCode & "            btnAdd_Click" & vbCrLf
    sCode = sCode & "        End If" & vbCrLf
    sCode = sCode & "        ' If the user clicked No, or if the Add failed, the text box will still have text." & vbCrLf
    sCode = sCode & "        ' Do not proceed with the replacement if the input field is not clear." & vbCrLf
    sCode = sCode & "        If Trim(Me.txtFind.Value) <> vbNullString Then Exit Sub" & vbCrLf
    sCode = sCode & "    End If" & vbCrLf & vbCrLf
    ' *** NEW SIMPLIFIED LOGIC BLOCK END ***
    
    sCode = sCode & "    Dim findReplacePairs As Variant" & vbCrLf
    sCode = sCode & "    Dim matchCase As Boolean, matchEntire As Boolean" & vbCrLf
    sCode = sCode & "    Dim ws As Worksheet, targetRange As Range" & vbCrLf
    sCode = sCode & "    Dim dataArray As Variant" & vbCrLf
    sCode = sCode & "    Dim i As Long, r As Long, c As Long" & vbCrLf
    sCode = sCode & "    Dim replacementCount As Long, sheetsAffected As Long" & vbCrLf
    sCode = sCode & "    Dim sheetsToProcess As Collection" & vbCrLf
    sCode = sCode & "    Dim compareMethod As VbCompareMethod" & vbCrLf & vbCrLf
    
    sCode = sCode & "    ' --- Validation and Setup ---" & vbCrLf
    sCode = sCode & "    If Me.lstPairs.ListCount = 0 Then" & vbCrLf
    sCode = sCode & "        MsgBox ""The replacement list is empty. Please add at least one pair."", vbExclamation, ""List Empty""" & vbCrLf
    sCode = sCode & "        Exit Sub" & vbCrLf
    sCode = sCode & "    End If" & vbCrLf
    sCode = sCode & "    findReplacePairs = Me.lstPairs.List" & vbCrLf
    sCode = sCode & "    matchCase = Me.chkMatchCase.Value" & vbCrLf
    sCode = sCode & "    matchEntire = Me.chkMatchEntire.Value" & vbCrLf
    sCode = sCode & "    compareMethod = IIf(matchCase, vbBinaryCompare, vbTextCompare)" & vbCrLf & vbCrLf
    
    sCode = sCode & "    ' --- Determine Scope: Which sheets to process ---" & vbCrLf
    sCode = sCode & "    Set sheetsToProcess = New Collection" & vbCrLf
    sCode = sCode & "    If Me.optSelection.Value Then" & vbCrLf
    sCode = sCode & "        If TypeName(Selection) <> ""Range"" Then" & vbCrLf
    sCode = sCode & "            MsgBox ""Please select a valid range."", vbExclamation, ""Invalid Selection""" & vbCrLf
    sCode = sCode & "            Exit Sub" & vbCrLf
    sCode = sCode & "        End If" & vbCrLf
    sCode = sCode & "        On Error Resume Next" & vbCrLf
    sCode = sCode & "        If TypeName(ActiveSheet) = ""Worksheet"" Then sheetsToProcess.Add ActiveSheet" & vbCrLf
    sCode = sCode & "        On Error GoTo 0" & vbCrLf
    sCode = sCode & "    ElseIf Me.optActiveSheet.Value Then" & vbCrLf
    sCode = sCode & "        If TypeName(ActiveSheet) = ""Worksheet"" Then sheetsToProcess.Add ActiveSheet" & vbCrLf
    sCode = sCode & "    Else ' Workbook" & vbCrLf
    sCode = sCode & "        For Each ws In ThisWorkbook.Worksheets" & vbCrLf
    sCode = sCode & "            sheetsToProcess.Add ws" & vbCrLf
    sCode = sCode & "        Next ws" & vbCrLf
    sCode = sCode & "    End If" & vbCrLf
    sCode = sCode & "    If sheetsToProcess.Count = 0 Then Exit Sub" & vbCrLf & vbCrLf

    sCode = sCode & "    ' --- Performance Optimization ---" & vbCrLf
    sCode = sCode & "    Application.ScreenUpdating = False" & vbCrLf
    sCode = sCode & "    Application.EnableEvents = False" & vbCrLf
    sCode = sCode & "    Application.Calculation = xlCalculationManual" & vbCrLf
    sCode = sCode & "    Me.Hide" & vbCrLf & vbCrLf

    sCode = sCode & "    ' --- Main Processing Loop ---" & vbCrLf
    sCode = sCode & "    For Each ws In sheetsToProcess" & vbCrLf
    sCode = sCode & "        Dim sheetHadReplacements As Boolean" & vbCrLf
    sCode = sCode & "        Dim originalValue As String, currentValue As String" & vbCrLf
    sCode = sCode & "        On Error Resume Next" & vbCrLf
    sCode = sCode & "        If Me.optSelection.Value Then" & vbCrLf
    sCode = sCode & "            Set targetRange = Intersect(Selection, ws.UsedRange)" & vbCrLf
    sCode = sCode & "        Else" & vbCrLf
    sCode = sCode & "            Set targetRange = ws.UsedRange" & vbCrLf
    sCode = sCode & "        End If" & vbCrLf
    sCode = sCode & "        If Err.Number <> 0 Or targetRange Is Nothing Then GoTo NextSheet" & vbCrLf
    sCode = sCode & "        Err.Clear" & vbCrLf
    sCode = sCode & "        On Error GoTo 0" & vbCrLf
    sCode = sCode & "        If targetRange.Cells.Count > 1 Then" & vbCrLf
    sCode = sCode & "             dataArray = targetRange.Value2" & vbCrLf
    sCode = sCode & "        Else" & vbCrLf
    sCode = sCode & "             ReDim dataArray(1 To 1, 1 To 1): dataArray(1, 1) = targetRange.Value2" & vbCrLf
    sCode = sCode & "        End If" & vbCrLf & vbCrLf

    sCode = sCode & "        ' --- Perform replacements in the array (very fast) ---" & vbCrLf
    sCode = sCode & "        For r = 1 To UBound(dataArray, 1)" & vbCrLf
    sCode = sCode & "            For c = 1 To UBound(dataArray, 2)" & vbCrLf
    sCode = sCode & "                If VarType(dataArray(r, c)) = vbString Then" & vbCrLf
    sCode = sCode & "                    originalValue = dataArray(r, c)" & vbCrLf
    sCode = sCode & "                    currentValue = originalValue" & vbCrLf & vbCrLf
    sCode = sCode & "                    For i = LBound(findReplacePairs, 1) To UBound(findReplacePairs, 1)" & vbCrLf
    sCode = sCode & "                        If matchEntire Then" & vbCrLf
    sCode = sCode & "                            If StrComp(currentValue, CStr(findReplacePairs(i, 0)), compareMethod) = 0 Then" & vbCrLf
    sCode = sCode & "                                currentValue = CStr(findReplacePairs(i, 1))" & vbCrLf
    sCode = sCode & "                            End If" & vbCrLf
    sCode = sCode & "                        Else" & vbCrLf
    sCode = sCode & "                            currentValue = Replace(currentValue, CStr(findReplacePairs(i, 0)), CStr(findReplacePairs(i, 1)), 1, -1, compareMethod)" & vbCrLf
    sCode = sCode & "                        End If" & vbCrLf
    sCode = sCode & "                    Next i" & vbCrLf & vbCrLf
    sCode = sCode & "                    If originalValue <> currentValue Then" & vbCrLf
    sCode = sCode & "                        dataArray(r, c) = currentValue" & vbCrLf
    sCode = sCode & "                        replacementCount = replacementCount + 1" & vbCrLf
    sCode = sCode & "                        sheetHadReplacements = True" & vbCrLf
    sCode = sCode & "                    End If" & vbCrLf
    sCode = sCode & "                End If" & vbCrLf
    sCode = sCode & "            Next c" & vbCrLf
    sCode = sCode & "        Next r" & vbCrLf & vbCrLf
    
    sCode = sCode & "        ' --- Write the modified array back to the sheet in one go ---" & vbCrLf
    sCode = sCode & "        If sheetHadReplacements Then" & vbCrLf
    sCode = sCode & "            On Error Resume Next ' Handle protected sheet write error" & vbCrLf
    sCode = sCode & "            targetRange.Value2 = dataArray" & vbCrLf
    sCode = sCode & "            If Err.Number = 0 Then sheetsAffected = sheetsAffected + 1" & vbCrLf
    sCode = sCode & "            On Error GoTo 0" & vbCrLf
    sCode = sCode & "        End If" & vbCrLf
    sCode = sCode & "NextSheet:" & vbCrLf
    sCode = sCode & "        Set targetRange = Nothing" & vbCrLf
    sCode = sCode & "    Next ws" & vbCrLf & vbCrLf
    
    sCode = sCode & "    ' --- Cleanup and Final Report ---" & vbCrLf
    sCode = sCode & "    Application.Calculation = xlCalculationAutomatic" & vbCrLf
    sCode = sCode & "    Application.EnableEvents = True" & vbCrLf
    sCode = sCode & "    Application.ScreenUpdating = True" & vbCrLf
    sCode = sCode & "    Me.Show" & vbCrLf
    sCode = sCode & "    MsgBox ""Replacement complete."" & vbCrLf & vbCrLf & ""Total replacements made: "" & replacementCount & vbCrLf & ""Sheets affected: "" & sheetsAffected, vbInformation, ""Operation Finished""" & vbCrLf
    sCode = sCode & "End Sub"

    ' Write the code to the form's code module
    codeModule.AddFromString sCode

    CreateMultiReplaceForm = True ' Success
    Exit Function

ErrorHandler:
    Dim errorMsg As String
    errorMsg = "An error occurred while trying to create the UserForm." & vbCrLf & vbCrLf & _
               "This can be caused by either of the following:" & vbCrLf & vbCrLf & _
               "1. 'Trust access to the VBA project object model' is not enabled in the Trust Center." & vbCrLf & _
               "2. The workbook's VBA project is password protected." & vbCrLf & vbCrLf & _
               "Please check these settings and try again."
    MsgBox errorMsg, vbCritical, "Form Creation Failed"
    
    ' Clean up a partially created form if it exists
    On Error Resume Next
    vbProj.VBComponents.Remove uf
    On Error GoTo 0
    CreateMultiReplaceForm = False ' Failure
End Function
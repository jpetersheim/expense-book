Attribute VB_Name = "mUniversals"
Public Sub progressbar(current, Max)
    'Call this at the end of your loop to track progress in it
    'Current = current position in loop (x in "x out of y")
    'Max = max number that the loop will go to (y in "x out of y")
    
    'Show user form
    If IsLoaded("frmProgress") = False Then
        frmProgress.Show
    End If
    
    'Math from variables
    pctCompl = current / Max * 100
    pctCompl = Round(pctCompl)

    'Update user form
    frmProgress.lblComplete.Caption = pctCompl & "%"
    frmProgress.lblProgBar.Width = pctCompl * 2
    
    If pctCompl = 100 Then
        frmProgress.lblWorking.Caption = "Complete!"
        MsgBox "Complete!", vbExclamation
        Unload frmProgress
    End If

    DoEvents
    
End Sub

Public Function IsLoaded(formName As String) As Boolean
'Checks if a form is loaded and returns true or false
Dim frm As Object
For Each frm In VBA.UserForms
    If frm.Name = formName Then
        IsLoaded = True
        Exit Function
    End If
Next frm
IsLoaded = False
End Function

Public Function dlgCheck(selectedCount As Integer, minItems As Integer, maxItems As Integer)
'Checks import file dialog, returns True if passes, False if fails
'selectedCount = # of selected files
'minItems = min # of files that can be selected
'maxItems = max # of files that can be selected

dlgCheck = True

'Checks if you selected greater than min
    If selectedCount < minItems Then
        MsgBox ("Incorrect amount of files selected. Please select at least " & minItems & " files.")
        dlgCheck = False
    End If

'Repeats # of selected items and confirms
    If selectedCount >= minItems And selectedCount <= maxItems Then
        y = MsgBox("You have selected " & selectedCount & " files.  Is this correct?.", 4132, "Confirm Selections")
        If y = vbNo Then
            dlgCheck = False
        End If
    End If
    
'Error if more than max items selected
    If selectedCount > maxItems Then
        MsgBox ("You have selected more than " & maxItems & " files.  Please click the button again and select " & maxItems & " or fewer files.")
        dlgCheck = False
    End If
End Function

Public Function GetUnique(uniqueRange As Range)
Dim x
Dim objDict As Object
Dim lngRow As Long

    Set objDict = CreateObject("Scripting.Dictionary")
    x = Application.Transpose(uniqueRange)

    If uniqueRange.Count <= 1 Then
        GetUnique = Array()
        Exit Function
    End If

    For lngRow = 1 To UBound(x, 1)
        objDict(x(lngRow)) = 1
    Next

    GetUnique = objDict.Keys

End Function

Public Function GetUniqueIf(uniqueRange As Range, conditionRange As Range, condition As Variant)
Dim x
Dim objDict As Object
Dim lngRow As Long

    Set objDict = CreateObject("Scripting.Dictionary")
    x = Application.Transpose(uniqueRange)
    Z = Application.Transpose(conditionRange)
    
    If uniqueRange.Count <= 1 Then
        GetUniqueIf = Array()
        Exit Function
    End If
    
    If UBound(x, 1) <> UBound(Z, 1) Then
        MsgBox ("Ranges must be same length.")
        End
    End If

    For lngRow = 1 To UBound(x, 1)
        If Z(lngRow) = condition Then
            objDict(x(lngRow)) = 1
        End If
    Next

    GetUniqueIf = objDict.Keys

End Function

Public Function GetEmptyRow(ws As Worksheet, column As Integer, startRow As Integer)
    Z = startRow
    Do Until ws.Cells(Z, column) = ""
        Z = Z + 1
    Loop
    
    GetEmptyRow = Z
End Function

Public Function ArrayLen(arr1 As Variant) As Integer
    ArrayLen = UBound(arr1) - LBound(arr1)
End Function

Public Function Max(x, y As Variant) As Variant
  Max = IIf(x > y, x, y)
End Function

Public Function Min(x, y As Variant) As Variant
   Min = IIf(x < y, x, y)
End Function


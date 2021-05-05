Attribute VB_Name = "mTasks"
Public Sub GetCategs()
    Dim categArr As Variant
    Dim currCategArr As Variant
    Dim addCateg As Boolean
    
    maxCateg = 20
    finalCategRow = 10 + maxCateg
    
    Application.ScreenUpdating = False

    Set expListBook = ActiveWorkbook
    Set expListSheet = expListBook.Sheets("Expense List")
    
    expListSheet.Activate
    
    lastRow = expListSheet.Range("A65536").End(xlUp).Row
    categArr = GetUnique(expListSheet.Range(Cells(3, 6), Cells(lastRow, 6)))
    
    expListBook.Sheets("Main Tab").Activate
    expListBook.Sheets("Main Tab").Range(Cells(11, 6), Cells(finalCategRow, 6)).ClearContents
    
    If UBound(categArr) >= 20 Then
        overCategs = UBound(categArr) - 20
        For j = 20 To UBound(categArr)
            If j = UBound(categArr) Then
                msgCategs = msgCategs & categArr(j)
            Else
                msgCategs = msgCategs & categArr(j) & ", "
            End If
        Next j
    
    MsgBox "You have " & UBound(categArr) + 1 & " categories. This workbook is currently limited to 20 categories." & vbCrLf & "Categories " & msgCategs & " are not listed to use."
    End If
    
    expListBook.Sheets("Main Tab").Cells(11, 6).Resize(Min(UBound(categArr) + 1, maxCateg)) = WorksheetFunction.Transpose(categArr)
    
'    emptyCategRow = GetEmptyRow(expListBook.Sheets("Main Tab"), 6, 11)
'
'    If emptyCategRow >= (finalCategRow + 2) Then
'        arrCats = GetUnique(Sheets("Main Tab").Range("F11:F" & finalCategRow))
'    ElseIf emptyCategRow = 11 Then
'        arrCats = GetUnique(Sheets("Main Tab").Range("F11:F12"))
'    Else
'        arrCats = GetUnique(Sheets("Main Tab").Range("F11", Range("F11").End(xlDown)))
'    End If
    
    Sheets("Working Sheet").Visible = True
    Sheets("Working Sheet").Activate
    Sheets("Working Sheet").Range("D5", Range("D5").End(xlDown)).Clear
    Sheets("Working Sheet").Cells(5, 4).Resize(Min(UBound(categArr) + 1, maxCateg)) = WorksheetFunction.Transpose(categArr)
    
    lastRow = GetEmptyRow(Sheets("Working Sheet"), 4, 3) - 1
    ActiveWorkbook.Names("Cat_List").RefersTo = Sheets("Working Sheet").Range(Cells(3, 4), Cells(lastRow, 4))
    Sheets("Working Sheet").Visible = False
    
    Sheets("Main Tab").Activate
    Application.ScreenUpdating = True
    
End Sub

Public Sub ExportSettings()
    Dim sFolder As String
    Dim wsVar As Worksheet
    
    Set wsVar = ActiveWorkbook.Sheets("Account Variables")
    
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a folder to save your bank data import settings"
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then ' if a file was chosen
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        strFileName = "ExpenseBook_DataImportSettings_" & Format(Now(), "DDMMMYYYY")
        strFullName = sFolder & "\" & strFileName
        
        wsVar.Visible = xlSheetVisible
        wsVar.Copy
        ActiveWorkbook.SaveAs Filename:=strFullName, FileFormat:=xlCSV, CreateBackup:=True
        ActiveWorkbook.Close
        
        wsVar.Visible = xlSheetHidden
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        MsgBox "Data import settings saved as a CSV file under " & strFullName
    Else
        MsgBox "No folder selected. Nothing happened."
    End If
    
End Sub

Public Sub ImportSettings()
    Dim sFile As String
    Dim wsVar As Worksheet
    Dim newWB As Workbook
    
    Set wsVar = ActiveWorkbook.Sheets("Account Variables")
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "CSV", "*.CSV", 1
        .Title = "Choose data import settings file"
        .AllowMultiSelect = False
        
        If .Show = True Then
            sFile = .SelectedItems(1)
        End If
    End With
    
    If sFile = "" Then
        MsgBox "No file selected."
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Workbooks.Open (sFile)
    Set newWB = ActiveWorkbook
    
    If newWB.Sheets(1).Cells(1, 15).Value = wsVar.Cells(1, 15).Value Then
        newWB.Sheets(1).Cells.Copy (wsVar.Cells(1, 1))
    Else
        MsgBox ("Those settings are from a previous version. Some settings may have changed since then. Please set up your accounts again.")
    End If
    
    newWB.Close
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Your data import settings have been applied."
End Sub

Public Function ShowNewAccForm(Optional Modal As Boolean = False)
    'API to display a userform in THIS workbook, from another workbook

     On Error Resume Next
     frmNewAcc.Show IIf(Modal, vbModal, vbModeless)
     ShowNewAccForm = (Err.Number = 0)
End Function


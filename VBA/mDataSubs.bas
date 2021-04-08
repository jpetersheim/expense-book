Attribute VB_Name = "mDataSubs"
Public firstEmptyExpRow As Integer
Public currTrans As Integer
Public compName As String, accType As String, accNick As String
Public importIncome As Boolean
Public autoPick As Boolean
Public lastFilesLocation As String
Public compArr(), accArr() As Variant
Public sourcesArr() As Variant, datesArr() As Variant, filesArr() As Variant
Public expListBook As Workbook, expListSheet As Worksheet, setupSheet As Worksheet, varSheet As Worksheet
Public expFile As Workbook, expSheet As Worksheet, expFileName As String
Public listTransCol, listDateCol, listTimeCol, listAmtCol, listCDescCol, listTypeCol, listCCatCol, listPCatCol, listCompCol, listPDescCol, listSrcCol, listAddDate As Integer

Public Sub ImportData()

    Dim dlgOpen As FileDialog
    Dim expFileNameL As String
    
    'Define names
    Set expListBook = ActiveWorkbook
    Set expListSheet = expListBook.Sheets("Expense List")
    Set setupSheet = expListBook.Sheets("Main Tab")
    Set varSheet = expListBook.Sheets("Account Variables")
    
    'Get updated company and account lists
    varSheet.Visible = True
    varSheet.Activate
    lastVarSheetRow = GetEmptyRow(varSheet, 1, 2) - 1
    If lastVarSheetRow = 2 Then
        compArr = Array(varSheet.Cells(2, 1).Value)
        accArr = Array(varSheet.Cells(2, 2).Value)
    Else
        compArr = GetUnique(varSheet.Range(Cells(2, 1), Cells(lastVarSheetRow, 1)))
        accArr = GetUniqueIf(varSheet.Range(Cells(2, 2), Cells(lastVarSheetRow, 2)), _
            varSheet.Range(Cells(2, 1), Cells(lastVarSheetRow, 1)), compName)
    End If

    'Show the form for information input
    frmSelectImport.Show
    
    'Import data from the form
    compName = frmSelectImport.compName
    accType = frmSelectImport.accType
    accNick = frmSelectImport.accNick
    importIncome = frmSelectImport.importIncome
    autoPick = frmSelectImport.autoPick
    
    'Set data columns in expense book
    listTransCol = 1
    listDateCol = 2
    listAmtCol = 3
    listCDescCol = 4
    listCCatCol = 5
    listPCatCol = 6
    listCompCol = 7
    listLocCol = 8
    listPDescCol = 9
    listSrcCol = 10
    listAddDate = 11
    
    'Number of forms that failed to import
    failedExp = 0
    
    'Dialog to select files
    Set dlgOpen = Application.FileDialog(FileDialogType:=msoFileDialogOpen)
    With dlgOpen
        If autoPick = True Then
            .Title = "Choose your activity or transaction files to be imported"
        Else
            .Title = "Choose your activity or transaction files for " & compName & " " & accType
        End If
        .Filters.Add Description:="Comma delimited files (*.CSV)", Extensions:="*.csv"
        .InitialFileName = lastFilesLocation
        .ButtonName = "Select"
        .AllowMultiSelect = True
        .Show
    End With
    
    'Make sure selected items fit the range
    fileCheck = dlgCheck(dlgOpen.SelectedItems.Count, 1, 40)
    
    If fileCheck = False Then
        varSheet.Visible = xlSheetHidden
        setupSheet.Activate
        End
    End If
    
    lastFilesLocation = Left(dlgOpen.SelectedItems(1), InStrRev(dlgOpen.SelectedItems(1), "\"))
    
    'Open up each submission file selected
    For expForm = 1 To dlgOpen.SelectedItems.Count
        Application.ScreenUpdating = False
        
        expFileName = Right(dlgOpen.SelectedItems(expForm), Len(dlgOpen.SelectedItems(expForm)) - InStrRev(dlgOpen.SelectedItems(expForm), "\"))
        expFileNameL = LCase(expFileName)
        
        expListSheet.Activate
        lastRow = expListSheet.Range("K65536").End(xlUp).Row
        
        'Determine if this file has been uploaded before
        currFilesArr = GetUnique(expListSheet.Range(Cells(3, 12), Cells(lastRow, 12)))
        
        For impFile = 0 To ArrayLen(currFilesArr)
            If currFilesArr(impFile) = expFileName Then
                MsgBox ("Duplicate filename: " & expFileName & vbNewLine & "The file is skipped. Please rename the file and try again.")
                failedExp = failedExp + 1
                GoTo NextExpForm
            End If
        Next impFile
        
        setupSheet.Cells(7, 3).Value = Now
        setupSheet.Cells(8, 3).Value = expFileName
        
        If autoPick = True Then
            For Each i In compArr
                newi = LCase(i)
                If InStr(expFileNameL, newi) <> 0 Then
                    compName = i
                End If
            Next i
            
            'Update acc list for chosen company
            varSheet.Activate
            accArr = GetUniqueIf(varSheet.Range(Cells(2, 2), Cells(lastVarSheetRow, 2)), _
                varSheet.Range(Cells(2, 1), Cells(lastVarSheetRow, 1)), compName)
            
            For Each i In accArr
                newi = LCase(i)
                If InStr(expFileNameL, newi) <> 0 Then
                    accType = i
                End If
            Next i
            
            If compName = "" Then
                MsgBox "Could not identify the company in the file name: " & expFileName
                failedExp = failedExp + 1
                GoTo NextExpForm
            ElseIf accType = "" Then
                MsgBox "The account type for " & expFileName & " does not match any added accounts. Please add the account and try again."
                failedExp = failedExp + 1
                GoTo NextExpForm
            End If
            
            accNick = compName & " " & accType
        End If
        
        'Find row to start entering data
        firstEmptyExpRow = GetEmptyRow(expListSheet, 11, 3)
        
        expListSheet.Activate
        currTrans = Application.WorksheetFunction.Max(expListSheet.Range(Cells(4, 1), Cells(firstEmptyExpRow, 1)))
        
        'Open each workbook
        Workbooks.OpenText dlgOpen.SelectedItems.Item(expForm)
        Set expFile = ActiveWorkbook
        Set expSheet = ActiveSheet
    
        'Determine how to import data based on company and account
        Call ImportTransactions(compName, accType)
        
        expFile.Close SaveChanges:=False
    
        Call progressbar(expForm - failedExp, dlgOpen.SelectedItems.Count - failedExp)
        
NextExpForm:
    Next expForm
    
    Call SetTransIDs
    
    Unload frmSelectImport
    'varSheet.Visible = False
    expListSheet.Activate
    
End Sub

Public Sub ImportTransactions(company As String, account As String)
    Dim expDateCol As Integer
    Dim expDescCol As Integer
    Dim expAmtCol As Integer
    Dim expCatCol As Integer
    Dim hasTCAT As Boolean
    Dim multiTransCols As Boolean
    Dim negTransactions As Boolean
    Dim varSheet As Worksheet
    
    'Get variables for selected company and account
    Set varSheet = expListBook.Sheets("Account Variables")
    accVarRow = 1
    Do Until (varSheet.Cells(accVarRow, 1).Value = company And varSheet.Cells(accVarRow, 2).Value = account) Or accVarRow = 1001
        If accVarRow = 1000 Then
            MsgBox "The company: " & company & " and account type: " & account & " have not been entered into this workbook yet. Moving on to next file."
            Exit Sub
        End If
                  
        accVarRow = accVarRow + 1
    Loop
    
    'Set Variables
    negTransactions = CBool(varSheet.Cells(accVarRow, 3).Value)
    expDateCol = Range(varSheet.Cells(accVarRow, 4).Value & 1).column
    expDescCol = Range(varSheet.Cells(accVarRow, 5).Value & 1).column
    expAmtCol = Range(varSheet.Cells(accVarRow, 6).Value & 1).column
    expCatCol = Range(varSheet.Cells(accVarRow, 7).Value & 1).column
    firstRow = varSheet.Cells(accVarRow, 8).Value
    multiTransCols = CBool(varSheet.Cells(accVarRow, 9).Value)
    withdrawalcol = Range(varSheet.Cells(accVarRow, 10).Value & 1).column
    depositCol = Range(varSheet.Cells(accVarRow, 11).Value & 1).column
        
    lastRow = expSheet.Range("A65536").End(xlUp).Row
    tempError = 0
    
    For trans = firstRow To lastRow
        tdate = expSheet.Cells(trans, expDateCol).Value
        tdesc = expSheet.Cells(trans, expDescCol).Value
        
        'Is the trans/income in one column, or two, like PNC?
        If multiTransCols = True Then
            If importIncome = False And expSheet.Cells(trans, depositCol) <> "" Then
                tempError = tempError + 1
                GoTo NextTrans
            ElseIf importIncome = True And expSheet.Cells(trans, depositCol) <> "" Then
                tamt = expSheet.Cells(trans, depositCol)
            Else
                tamt = expSheet.Cells(trans, withdrawalcol)
            End If
        Else
            tamt = expSheet.Cells(trans, expAmtCol).Value
        End If
        
        'Does this institution supply categories?
        If hasTCAT = False Then
            tcat = "N/A"
        Else
            tcat = expSheet.Cells(trans, expCatCol).Value
        End If
        
        'Are transactions represented by negative or positive values?
        If negTransactions = True And importIncome = False And tamt > 0 Then
            tempError = tempError + 1
            GoTo NextTrans
        End If
        
        If negTransactions = False And importIncome = False And tamt < 0 Then
            tempError = tempError + 1
            GoTo NextTrans
        End If
        
        'Add the details to the expense sheet
        tamt = Abs(tamt)
        
        'First empty row on expense sheet + (the transaction number - first row on the sheet - any income values that we arent recording = row number on imported data)
        'expListSheet.Cells(firstEmptyExpRow + trans - firstRow - tempError, listTransCol).Value = currTrans + trans - firstRow + 1 - tempError
        expListSheet.Cells(firstEmptyExpRow + trans - firstRow - tempError, listDateCol).Value = Format(tdate, "dd-mmm-yyyy")
        expListSheet.Cells(firstEmptyExpRow + trans - firstRow - tempError, listCDescCol).Value = tdesc
        expListSheet.Cells(firstEmptyExpRow + trans - firstRow - tempError, listAmtCol).Value = tamt
        expListSheet.Cells(firstEmptyExpRow + trans - firstRow - tempError, listCCatCol).Value = tcat
        expListSheet.Cells(firstEmptyExpRow + trans - firstRow - tempError, listSrcCol).Value = accNick
        expListSheet.Cells(firstEmptyExpRow + trans - firstRow - tempError, listAddDate).Value = Format(Date, "dd-mmm-yyyy")
        expListSheet.Cells(firstEmptyExpRow + trans - firstRow - tempError, (listAddDate + 1)).Value = expFileName
        
NextTrans:
    Next trans
    
End Sub

Public Sub RemoveImports()

    Application.ScreenUpdating = False

    Set expListBook = ActiveWorkbook
    Set expListSheet = expListBook.Sheets("Expense List")
    
    expListSheet.Activate
    
    frmDeleteData.Show

End Sub

Public Sub SetupImports()
    frmCurrAccs.Show
End Sub

Public Sub SetTransIDs()
    
    Set expListBook = ActiveWorkbook
    Set expListSheet = expListBook.Sheets("Expense List")
    Application.ScreenUpdating = False
    
    prevSheet = ActiveSheet.Name
    expListSheet.Activate

    lastRow = GetEmptyRow(expListSheet, 11, 3) - 1
    
    For t = 3 To lastRow
        expListSheet.Cells(t, 1).Value = t - 2
    Next t
    
    If lastRow = 2 Then
        ActiveWorkbook.Names("ExpenseList").RefersTo = expListSheet.Range(Cells(2, 1), Cells(lastRow, 14))
        lastRow = 4
    Else
        expListSheet.Range(Cells(3, 13), Cells(lastRow, 13)).Formula = "=TEXT(B3,""mmmm"")"
        expListSheet.Range(Cells(3, 14), Cells(lastRow, 14)).Formula = "=YEAR(B3)"
        
        ActiveWorkbook.Names("ExpenseList").RefersTo = expListSheet.Range(Cells(2, 1), Cells(lastRow, 14))
    End If
    
    arrMonths = GetUnique(expListSheet.Range(Cells(3, 13), Cells(lastRow, 13)))
    arrYears = GetUnique(expListSheet.Range(Cells(3, 14), Cells(lastRow, 14)))
    
    ActiveWorkbook.Sheets("Working Sheet").Activate
    Sheets("Working Sheet").Range("B5", Range("B5").End(xlDown)).Clear
    Sheets("Working Sheet").Cells(5, 2).Resize(UBound(arrMonths) + 1) = WorksheetFunction.Transpose(arrMonths)
    
    Sheets("Working Sheet").Range("C5", Range("C5").End(xlDown)).Clear
    Sheets("Working Sheet").Cells(5, 3).Resize(UBound(arrYears) + 1) = WorksheetFunction.Transpose(arrYears)
    
    ActiveWorkbook.Worksheets("Working Sheet").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Working Sheet").Sort.SortFields.Add2 Key:=Range("B5", Range("B5").End(xlDown)), _
    SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "January,February,March,April,May,June,July,August,September,October,November,December" _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Working Sheet").Sort
        .SetRange Range("B5", Range("B5").End(xlDown))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveWorkbook.Worksheets("Working Sheet").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Working Sheet").Sort.SortFields.Add2 Key:=Range("C5", Range("C5").End(xlDown)), _
    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Working Sheet").Sort
        .SetRange Range("C5", Range("C5").End(xlDown))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    lastRow = GetEmptyRow(Sheets("Working Sheet"), 2, 3) - 1
    ActiveWorkbook.Names("Month_List").RefersTo = Sheets("Working Sheet").Range(Cells(3, 2), Cells(lastRow, 2))
    lastRow = GetEmptyRow(Sheets("Working Sheet"), 3, 3) - 1
    ActiveWorkbook.Names("Year_List").RefersTo = Sheets("Working Sheet").Range(Cells(3, 3), Cells(lastRow, 3))
    
    Sheets(prevSheet).Activate
    Application.ScreenUpdating = True
    
End Sub


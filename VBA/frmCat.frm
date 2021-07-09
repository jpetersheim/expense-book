VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCat 
   Caption         =   "Categorize Transcations"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14610
   OleObjectBlob   =   "frmCat.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lastExpenseRow As Integer
Public expListBook As Workbook
Public expListSheet As Worksheet
Public columnEmpty As Integer
Public acctCheck As String
Public totalAmt As Double

Private Sub cmbChsAcc_Change()

    Call ListTransactions(Me.txtSearch.Value, columnEmpty, Me.cmbChsAcc.Value, Me.cmbChsDate.Value)

End Sub

Private Sub cmbChsDate_Change()
    Call ListTransactions(Me.txtSearch.Value, columnEmpty, Me.cmbChsAcc.Value, Me.cmbChsDate.Value)
End Sub

Private Sub cmbEmptyCol_Change()
    Select Case cmbEmptyCol.Value
        Case "Category"
            columnEmpty = 6
        Case "Company"
            columnEmpty = 7
        Case "Location"
            columnEmpty = 8
        Case "Description"
            columnEmpty = 9
        Case ""
            columnEmpty = 0
    End Select
            
    Call ListTransactions(Me.txtSearch.Value, columnEmpty, Me.cmbChsAcc.Value, Me.cmbChsDate.Value)
            
End Sub

Private Sub cmdAutofill_Click()
    leaveEmpty = MsgBox("Would you like transactions with multiple previous categories to be left blank? If No then the first category listed in the Expense List will be used.", vbYesNo)
    
    With lsbTransactions
        For j = 0 To .ListCount - 1
            jRow = .List(j, 6) + 2
            Set Rng = Sheets("Expense List").Range("A" & (jRow) & ":L" & (jRow))
            uniqueName = .List(j, 2)
            
            Sheets("Expense List").Activate
            tempArr = GetUniqueIf(Sheets("Expense List").Range(Cells(3, 6), Cells(lastExpenseRow, 6)), _
                        Sheets("Expense List").Range(Cells(3, 4), Cells(lastExpenseRow, 4)), uniqueName)
                        
            ' used 2 because the empty value counts in the array, so if it has 2 previous itll be 3
            If UBound(tempArr) >= 2 And leaveEmpty = vbYes Then
                Rng.Cells(1, 6).Value = ""
            Else
                Rng.Cells(1, 6).Value = tempArr(0)
            End If
        Next j
    End With
    
    Call UpdateArrays
End Sub

Private Sub cmdCat_Click()
    With lsbTransactions
        For j = 0 To .ListCount - 1
            If .Selected(j) = True Then
                jRow = .List(j, 6) + 2
                Set Rng = Sheets("Expense List").Range("A" & (jRow) & ":L" & (jRow))
                Rng.Cells(1, 6).Value = Me.cmbCategory.Value
                Rng.Cells(1, 7).Value = Me.cmbCompany.Value
                Rng.Cells(1, 8).Value = Me.cmbLocation
                Rng.Cells(1, 9).Value = Me.txtDescription
            End If
        Next j
    End With
    
    Call UpdateArrays
End Sub

Private Sub cmdCatList_Click()
    With lsbTransactions
        For j = 0 To .ListCount - 1
            jRow = .List(j, 6) + 2
            Set Rng = Sheets("Expense List").Range("A" & (jRow) & ":L" & (jRow))
            Rng.Cells(1, 6).Value = Me.cmbCategory.Value
            Rng.Cells(1, 7).Value = Me.cmbCompany.Value
            Rng.Cells(1, 8).Value = Me.cmbLocation
            Rng.Cells(1, 9).Value = Me.txtDescription
        Next j
    End With
    
    Call UpdateArrays
End Sub

Private Sub cmdDelTrans_Click()
    tDelNum = 0
    DelYN = MsgBox("Are you sure you want to delete the selected transactions?", vbYesNo)
    If DelYN = vbYes Then
        With lsbTransactions
            For j = 0 To .ListCount - 1
                If .Selected(j) = True Then
                    jRow = .List(j, 6) + 2 - tDelNum
                    expListSheet.Rows(jRow).Delete
                    tDelNum = tDelNum + 1
                End If
            Next j
        End With
    End If
    
    Call SetTransIDs
    Call UpdateArrays
    Call ListTransactions
End Sub

Private Sub cmdDeselAll_Click()
    Call ListSelect(lsbTransactions, False)
End Sub

Private Sub cmdResetList_Click()
    Me.cmbEmptyCol.Value = ""
    columnEmpty = 0
    Me.txtSearch.Value = ""
    Me.cmbChsAcc.Value = ""
    Me.cmbChsDate.Value = ""
    Call ListTransactions
End Sub

Private Sub cmdRevert_Click()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    expListBook.Sheets("Backup Expense List").Visible = True
    expListBook.Sheets("Backup Expense List").Cells.Copy Destination:=expListSheet.Cells
    expListBook.Sheets("Backup Expense List").Visible = False
    expListBook.Sheets("Main Tab").Activate
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Call SetTransIDs
    Call GetCategs
    
    Unload Me
    
End Sub

Private Sub cmdSave_Click()
    Call SetTransIDs
    Call BackupExpenses
    Call GetCategs
End Sub

Private Sub cmdSaveClose_Click()
    Call SetTransIDs
    Call BackupExpenses
    Call GetCategs
    
    Unload Me
End Sub

Private Sub cmdSelAll_Click()
    Call ListSelect(lsbTransactions, True)
End Sub

Private Sub cmdSplitTrans_Click()
    tcount = 0
    With lsbTransactions
        For j = 0 To .ListCount - 1
            If .Selected(j) = True Then
                tcount = tcount + 1
            End If
        Next j
    End With
    
    If tcount > 1 Then
        MsgBox "Can't split more than 1 transaction at a time."
    ElseIf tcount = 0 Then
        MsgBox "You must select a transaction to split."
    End If
    
    transID = lsbTransactions.List(lsbTransactions.ListIndex, 6)
    totalAmt = lsbTransactions.List(lsbTransactions.ListIndex, 1)
    
    frmSplitTrans.lblTotalAmt = "$" & totalAmt
    frmSplitTrans.lblTrans2Amt = "$0.00"
    frmSplitTrans.txtTrans1Amt.Value = 0
    frmSplitTrans.Show
    
    If frmSplitTrans.cancelVar = True Then
        Exit Sub
    End If
    
    oneAmt = frmSplitTrans.txtTrans1Amt.Value
    twoAmtStr = frmSplitTrans.lblTrans2Amt
    twoAmt = Right(twoAmtStr, Len(twoAmtStr) - 1)
    
    expListSheet.Rows(transID + 2).Insert
    For colNum = 2 To 15
        expListSheet.Cells(transID + 2, colNum) = expListSheet.Cells(transID + 3, colNum)
        If colNum = 3 Then
            expListSheet.Cells(transID + 2, colNum) = oneAmt
            expListSheet.Cells(transID + 3, colNum) = twoAmt
        End If
            
        If colNum = 5 Then colNum = colNum + 1
    Next colNum
    
    Call SetTransIDs
    Call ListTransactions(Me.txtSearch.Value, columnEmpty, Me.cmbChsAcc.Value, Me.cmbChsDate.Value)
    
End Sub

Private Sub cmdUpdate_Click()
    Call ListTransactions(Me.txtSearch.Value, columnEmpty, Me.cmbChsAcc.Value, Me.cmbChsDate.Value)
End Sub

Private Sub cmdX_Click()
    Me.txtSearch.Value = ""
End Sub

Private Sub lsbTransactions_Change()
    tcount = 0
    With lsbTransactions
        For j = 0 To .ListCount - 1
            If .Selected(j) = True Then
                tcount = tcount + 1
            End If
        Next j
    
        Me.cmbCategory = ""
        Me.cmbCompany = ""
        Me.cmbLocation = ""
        Me.txtDescription = ""

        If tcount = 1 Then
            For j = 0 To .ListCount - 1
                If .Selected(j) = True Then
                    jRow = .List(j, 6) + 2
                    Set Rng = Sheets("Expense List").Range("A" & (jRow) & ":L" & (jRow))
                    Me.cmbCategory = Rng.Cells(1, 6).Value
                    Me.cmbCompany = Rng.Cells(1, 7).Value
                    Me.cmbLocation = Rng.Cells(1, 8).Value
                    Me.txtDescription = Rng.Cells(1, 9).Value
                End If
            Next j
        End If
    
    End With
    
    Me.lblSelNum = tcount
End Sub

Private Sub txtSearch_Change()

    Call ListTransactions(txtSearch.Value, columnEmpty, Me.cmbChsAcc.Value, Me.cmbChsDate.Value)

End Sub

Private Sub UserForm_Initialize()
    Dim colArray As Variant
    
    Set expListBook = ActiveWorkbook
    Set expListSheet = expListBook.Sheets("Expense List")

    colArray = Array("Category", "Company", "Location", "Description", "")
    cmbEmptyCol.List = colArray

    lastExpenseRow = GetEmptyRow(expListSheet, 11, 3) - 1
    
    expListSheet.Activate
    
    Call UpdateArrays
    Call BackupExpenses
    Call ListTransactions
    
    cmbEmptyCol.Value = "Category"
    
End Sub

Private Sub ListTransactions(Optional filter As String, Optional colEmpty As Integer, Optional accFilter As String, Optional dateFilter As String)
    Dim matchesFilter As Boolean
    Dim arrValues() As Variant
    lsbTransactions.Clear

    'LastRow = Sheets("Expense List").Range("A65536").End(xlUp).Row
    If filter = vbNullString And colEmpty = 0 And accFilter = "" And dateFilter = "" Then
    
        For j = 3 To lastExpenseRow
        
            Set Rng = Sheets("Expense List").Range("A" & j & ":L" & j)
            
            With Me.lsbTransactions
                .AddItem
                .List(c, 0) = Rng.Cells(1, 2).Value
                .List(c, 1) = Rng.Cells(1, 3).Value
                .List(c, 2) = Rng.Cells(1, 4).Value
                .List(c, 3) = Rng.Cells(1, 5).Value
                .List(c, 4) = Rng.Cells(1, 10).Value
                .List(c, 5) = Rng.Cells(1, 11).Value
                .List(c, 6) = Rng.Cells(1, 1).Value
                c = c + 1
            End With
            
        Next j
    
    Else
    
        For j = 3 To lastExpenseRow
            matchesFilter = False
            matchesAcct = False
            matchesEmpty = False
            matchesDate = False
            
            If filter = "" Then
                matchesFilter = True
            End If
            
            If colEmpty = 0 Then
                matchesEmpty = True
            End If
            
            If accFilter = "" Then
                matchesAcct = True
            End If
            
            If dateFilter = "" Then
                matchesDate = True
            End If
            
            Set Rng = Sheets("Expense List").Range("A" & j & ":L" & j)
            newFilter = UCase("*" & filter & "*")
                 
            arrValues = Rng.Value
            
            For k = LBound(arrValues, 2) To UBound(arrValues, 2)
                If UCase(arrValues(1, k)) Like newFilter Then
                    matchesFilter = True
                End If
                
                If colEmpty <> 0 Then
                    If arrValues(1, colEmpty) = "" Then
                        matchesEmpty = True
                    End If
                End If
                
                If arrValues(1, 10) = Me.cmbChsAcc.Value Then
                    matchesAcct = True
                End If
                
                If CStr(arrValues(1, 11)) = Me.cmbChsDate.Value Then
                    matchesDate = True
                End If
            
            Next k
            
            If matchesFilter = True And matchesAcct = True And matchesEmpty = True And matchesDate = True Then
                With Me.lsbTransactions
                    .AddItem
                    .List(c, 0) = Rng.Cells(1, 2).Value
                    .List(c, 1) = Rng.Cells(1, 3).Value
                    .List(c, 2) = Rng.Cells(1, 4).Value
                    .List(c, 3) = Rng.Cells(1, 5).Value
                    .List(c, 4) = Rng.Cells(1, 10).Value
                    .List(c, 5) = Rng.Cells(1, 11).Value
                    .List(c, 6) = Rng.Cells(1, 1).Value
                    c = c + 1
                End With
            End If
            
        Next j
        
    End If
    
    Me.lblListNum = CStr(Me.lsbTransactions.ListCount)
    Me.lblSelNum = "0"
    'lsbTransactions.ListIndex = 1
End Sub

Private Sub ListSelect(lsb, all As Boolean)
    If all = True Then
        With lsb
            For j = 0 To .ListCount - 1
                .Selected(j) = True
            Next j
        End With
    End If
    
    If all = False Then
        With lsb
            For j = 0 To .ListCount - 1
                .Selected(j) = False
            Next j
        End With
    End If
End Sub

Public Sub BackupExpenses()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Sheets("Backup Expense List").Visible = True
    expListBook.Sheets("Backup Expense List").Delete
    expListSheet.Copy After:=expListSheet
    Sheets("Expense List (2)").Name = "Backup Expense List"
    Sheets("Backup Expense List").Visible = False
    Sheets("Main Tab").Activate
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Call cmdRevert_Click
    End If
End Sub

Private Sub UpdateArrays()
    Application.ScreenUpdating = False
    expListSheet.Activate

    accArray = GetUnique(expListSheet.Range(Cells(3, 10), Cells(lastExpenseRow, 10)))
    Me.cmbChsAcc.List = accArray
    
    If accArray(0) = "Account" Then
        MsgBox "It doesn't appear any data has been imported yet."
        expListBook.Sheets("Main Tab").Activate
        End
    End If
    
    Me.cmbChsAcc.AddItem ""
    
    dateArray = GetUnique(expListSheet.Range(Cells(3, 11), Cells(lastExpenseRow, 11)))
    For j = LBound(dateArray) To UBound(dateArray)
        dateArray(j) = CDate(dateArray(j))
    Next j
    Me.cmbChsDate.List = dateArray
    Me.cmbChsDate.AddItem ""
    
    catArray = GetUnique(expListSheet.Range(Cells(3, 6), Cells(lastExpenseRow, 6)))
    Me.cmbCategory.List = catArray
    
    compArray = GetUnique(expListSheet.Range(Cells(3, 7), Cells(lastExpenseRow, 7)))
    Me.cmbCompany.List = compArray
    
    locArray = GetUnique(expListSheet.Range(Cells(3, 8), Cells(lastExpenseRow, 8)))
    Me.cmbLocation.List = locArray
    
    expListBook.Sheets("Main Tab").Activate
    Application.ScreenUpdating = True
End Sub

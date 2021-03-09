VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewAcc 
   Caption         =   "Add new account"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   OleObjectBlob   =   "frmNewAcc.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNewAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public accVarWs As Worksheet
Public rowModify As Integer

Private Sub cmdAdd_Click()

    If Me.txtDateCol.Value = "" Then
        MsgBox "Date column value can not be blank."
        Exit Sub
    End If
    
    If Me.txtAmtCol.Value = "" Then
        MsgBox "Amount column value can not be blank."
        Exit Sub
    End If
    
    If Me.togMultiCol = "True" And Me.txtWithdrawalCol.Value = "" Then
        MsgBox "If there are multiple amount columns, then the withdrawal column can not be blank."
        Exit Sub
    End If
    
    If Me.txtCompany.Value = "" Then
        MsgBox "Company name can not be blank."
        Exit Sub
    End If
    
    If Me.cmbAccountType.Value = "" Then
        MsgBox "Account type can not be blank."
        Exit Sub
    End If
    
    If Me.txtRowNum.Value = "" Then
        MsgBox "Row number can not be blank."
        Exit Sub
    End If

    'TODO CHECK IF ACCOUNT EXISTS AND ASK ABOUT OVERWRITING IF NEW
    'SAME FOR MODIFYING
    
    ActiveWorkbook.Sheets("Account Variables").Visible = xlSheetVisible

    Set accVarWs = ActiveWorkbook.Sheets("Account Variables")
    newAccRow = GetEmptyRow(accVarWs, 1, 2)
    
    For j = 2 To (newAccRow + 1)
        If accVarWs.Cells(j, 1) = Me.txtCompany.Value And accVarWs.Cells(j, 2) = Me.cmbAccountType.Value And rowModify = 0 Then
            overYN = MsgBox("Import settings for " & Me.txtCompany.Value & " " & Me.cmbAccountType.Value & " already exist. Would you like to overwrite these settings?", vbYesNo)
            If overYN = vbYes Then
                newAccRow = j
            Else
                Exit Sub
            End If
        ElseIf rowModify > 0 Then
            newAccRow = rowModify
            If accVarWs.Cells(j, 1) = Me.txtCompany.Value And accVarWs.Cells(j, 2) = Me.cmbAccountType.Value Then
                over2YN = MsgBox("Import settings for " & Me.txtCompany.Value & " " & Me.cmbAccountType.Value & " already exist. Are you sure you want to modify/overwite these settings?", vbYesNo)
                If over2YN = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    Next j
        
    accVarWs.Cells(newAccRow, 1).Value = Me.txtCompany.Value
    accVarWs.Cells(newAccRow, 2).Value = Me.cmbAccountType.Value
    accVarWs.Cells(newAccRow, 3).Value = Me.togNegative.Value
    
    accVarWs.Cells(newAccRow, 4).Value = UCase(Me.txtDateCol.Value)
    
    If Me.txtDescCol.Value = "" Then
        accVarWs.Cells(newAccRow, 5).Value = "ZZ"
    Else
        accVarWs.Cells(newAccRow, 5).Value = UCase(Me.txtDescCol.Value)
    End If
    
    accVarWs.Cells(newAccRow, 6).Value = UCase(Me.txtAmtCol.Value)
    
    If Me.txtCatCol.Value = "" Then
        accVarWs.Cells(newAccRow, 7).Value = "ZZ"
    Else
        accVarWs.Cells(newAccRow, 7).Value = UCase(Me.txtCatCol.Value)
    End If
    
    accVarWs.Cells(newAccRow, 8).Value = Me.txtRowNum.Value
    accVarWs.Cells(newAccRow, 9).Value = Me.togMultiCol.Value
    
    If Me.togMultiCol.Value = "False" Then
        accVarWs.Cells(newAccRow, 10).Value = "ZZ"
        accVarWs.Cells(newAccRow, 11).Value = "ZZ"
    Else
        accVarWs.Cells(newAccRow, 10).Value = UCase(Me.txtWithdrawalCol.Value)
        If Me.txtDepositsCol.Value = "" Then
            accVarWs.Cells(newAccRow, 11).Value = "ZZ"
        Else
            accVarWs.Cells(newAccRow, 11).Value = UCase(Me.txtDepositsCol.Value)
        End If
    End If
    
    accVarWs.Visible = xlSheetHidden
    
    Unload Me
    
End Sub

Private Sub togMultiCol_Change()
    If Me.togMultiCol.Value = True Then
        Me.togMultiCol.Caption = "Yes"
        Me.txtWithdrawalCol.Enabled = True
        Me.txtWithdrawalCol.BackColor = vbWhite
        Me.txtDepositsCol.Enabled = True
        Me.txtDepositsCol.BackColor = vbWhite
    ElseIf Me.togMultiCol.Value = False Then
        Me.togMultiCol.Caption = "No"
        Me.txtWithdrawalCol.Enabled = False
        Me.txtWithdrawalCol.BackColor = vbGrey
        Me.txtDepositsCol.Enabled = False
        Me.txtDepositsCol.BackColor = vbGrey
        Me.txtDepositsCol.Value = ""
        Me.txtWithdrawalCol.Value = ""
    End If
End Sub

Private Sub togNegative_Change()
    If Me.togNegative.Value = True Then
        Me.togNegative.Caption = "Negative"
    ElseIf Me.togNegative.Value = False Then
        Me.togNegative.Caption = "Positive"
    End If
End Sub

Private Sub txtDateCol_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtDateCol Like "*[0-9]*" Or Me.txtDateCol Like "*[@#$%*^&?()<>/\'""!]*" Then
        Me.txtDateCol.Value = ""
        MsgBox "Invalid entry. Please input a column letter.", vbInformation
        Cancel = True
    End If
End Sub

Private Sub txtDescCol_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtDescCol Like "*[0-9]*" Or Me.txtDescCol Like "*[@#$%*^&?()<>/\'""!]*" Then
        Me.txtDescCol.Value = ""
        MsgBox "Invalid entry. Please input a column letter.", vbInformation
        Cancel = True
    End If
End Sub

Private Sub txtAmtCol_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtAmtCol Like "*[0-9]*" Or Me.txtAmtCol Like "*[@#$%*^&?()<>/\'""!]*" Then
        Me.txtAmtCol.Value = ""
        MsgBox "Invalid entry. Please input a column letter.", vbInformation
        Cancel = True
    End If
End Sub

Private Sub txtCatCol_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtCatCol Like "*[0-9]*" Or Me.txtCatCol Like "*[@#$%*^&?()<>/\'""!]*" Then
        Me.txtCatCol.Value = ""
        MsgBox "Invalid entry. Please input a column letter.", vbInformation
        Cancel = True
    End If
End Sub

Private Sub txtWithdrawalCol_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtWithdrawalCol Like "*[0-9]*" Or Me.txtWithdrawalCol Like "*[@#$%*^&?()<>/\'""!]*" Then
        Me.txtWithdrawalCol.Value = ""
        MsgBox "Invalid entry. Please input a column letter.", vbInformation
        Cancel = True
    End If
End Sub

Private Sub txtDepositCol_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtDepositCol Like "*[0-9]*" Or Me.txtDepositCol Like "*[@#$%*^&?()<>/\'""!]*" Then
        Me.txtDepositCol.Value = ""
        MsgBox "Invalid entry. Please input a column letter.", vbInformation
        Cancel = True
    End If
End Sub

Private Sub txtRowNum_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtRowNum Like "*[0-9]*" Then
    Else
        Me.txtRowNum.Value = ""
        MsgBox "Invalid entry. Please input a row number.", vbInformation
        Cancel = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.cmbAccountType.AddItem "Checking"
    Me.cmbAccountType.AddItem "Credit"
    Me.cmbAccountType.AddItem "Saving"
    
    Set accVarWs = ActiveWorkbook.Sheets("Account Variables")
    
    rowModify = frmCurrAccs.rowModify
    
    If rowModify > 0 Then
        Me.txtCompany.Value = accVarWs.Cells(rowModify, 1).Value
        Me.cmbAccountType.Value = accVarWs.Cells(rowModify, 2).Value
        Me.togNegative.Value = accVarWs.Cells(rowModify, 3).Value
        Me.txtDateCol.Value = accVarWs.Cells(rowModify, 4).Value
        Me.txtDescCol.Value = accVarWs.Cells(rowModify, 5).Value
        Me.txtAmtCol.Value = accVarWs.Cells(rowModify, 6).Value
        Me.txtCatCol.Value = accVarWs.Cells(rowModify, 7).Value
        Me.txtRowNum.Value = accVarWs.Cells(rowModify, 8).Value
        Me.togMultiCol.Value = accVarWs.Cells(rowModify, 9).Value
        Me.txtWithdrawalCol.Value = accVarWs.Cells(rowModify, 10).Value
        Me.txtDepositsCol.Value = accVarWs.Cells(rowModify, 11).Value
    End If
        
End Sub

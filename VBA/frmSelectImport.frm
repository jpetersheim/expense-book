VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectImport 
   Caption         =   "Select Options"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSelectImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public importIncome As Boolean
Public autoPick As Boolean
Public accNick As String
Public compName As String
Public accType As String

Private Sub chkAutoPick_Change()
    If chkAutoPick.Value = True Then
        Me.cmbCompany.Enabled = False
        Me.cmbCompany.BackColor = vbGrey
        Me.cmbAccType.Enabled = False
        Me.cmbAccType.BackColor = vbGrey
        Me.txbNick.Enabled = False
        Me.txbNick.BackColor = vbGrey
        lblNickInfo.Visible = True
        
        Me.cmbCompany.Value = ""
        Me.cmbAccType.Value = ""
        Me.txbNick.Value = ""
    Else
        Me.cmbCompany.Enabled = True
        Me.cmbCompany.BackColor = vbWhite
        Me.cmbAccType.Enabled = True
        Me.cmbAccType.BackColor = vbWhite
        Me.txbNick.Enabled = True
        Me.txbNick.BackColor = vbWhite
        lblNickInfo.Visible = False
    End If
    
End Sub

Private Sub cmbAccType_Change()
    txbNick.Value = cmbCompany.Value & " " & cmbAccType.Value
End Sub

Private Sub cmbCompany_Change()
    cmbAccType.Value = ""
    
    txbNick.Value = cmbCompany.Value & " " & cmbAccType.Value
    
    lastVarSheetRow = GetEmptyRow(varSheet, 1, 2) - 1
    If lastVarSheetRow = 2 Then
        accArr = Array(varSheet.Cells(2, 2).Value)
    Else
        accArr = GetUniqueIf(varSheet.Range(Cells(2, 2), Cells(lastVarSheetRow, 2)), _
            varSheet.Range(Cells(2, 1), Cells(lastVarSheetRow, 1)), cmbCompany.Value)
    End If
    
    cmbAccType.List = accArr
End Sub

Private Sub UserForm_Initialize()
    Application.ScreenUpdating = False
    If compArr(0) = "Company Name" Then
        cmbCompany.List = Array("")
    Else
        cmbCompany.List = compArr
    End If
    cmbAccType.List = accArr
End Sub

Private Sub cmdImport_Click()
    If (cmbAccType.Value = "" Or cmbCompany.Value = "") And chkAutoPick.Value = False Then
        MsgBox "Please enter a value for the Company and Account."
        Exit Sub
    End If

    accType = cmbAccType.Value
    compName = cmbCompany.Value
    accNick = txbNick.Value
    importIncome = chkIncome.Value
    autoPick = chkAutoPick.Value
    
    If chkAutoPick.Value = True Then
        MsgBox "The macro determines the company and account by the file name." & vbCrLf & vbCrLf & "Please make sure the correctly spelled company name and account type are in the file name."
    End If
    
    frmSelectImport.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    varSheet.Visible = False
    setupSheet.Activate
    
    If CloseMode = vbFormControlMenu Then
        End
    End If
End Sub

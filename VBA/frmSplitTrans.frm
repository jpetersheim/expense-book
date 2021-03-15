VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplitTrans 
   Caption         =   "Set Transaction Amounts"
   ClientHeight    =   2385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3990
   OleObjectBlob   =   "frmSplitTrans.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplitTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cancelVar As Boolean

Private Sub cmdDone_Click()
    If Me.txtTrans1Amt.Value = 0 Then
        MsgBox "Can't split a transaction and have an amount of $0."
        Exit Sub
    End If
    
    cancelVar = False
    Me.Hide
    
End Sub

Private Sub txtTrans1Amt_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtTrans1Amt Like "*[0-9]*" Then
    Else
        Me.txtTrans1Amt.Value = ""
        MsgBox "Invalid entry. Please input a number.", vbInformation
        Cancel = True
    End If
End Sub

Private Sub txtTrans1Amt_AfterUpdate()
    If frmCat.totalAmt - txtTrans1Amt.Value <= 0 Then
        MsgBox "Value too large."
        txtTrans1Amt.Value = 0
    End If

    If txtTrans1Amt.Value <> "" Then
        Me.lblTrans2Amt = "$" & Format((frmCat.totalAmt - txtTrans1Amt.Value), "#0.00")
    End If
    
End Sub

Private Sub UserForm_Initialize()
    cancelVar = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    cancelVar = True
    Me.Hide
  End If
End Sub


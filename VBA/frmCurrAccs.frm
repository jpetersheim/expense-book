VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCurrAccs 
   Caption         =   "Setup Data Import"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3945
   OleObjectBlob   =   "frmCurrAccs.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCurrAccs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public accVarWs As Worksheet
Public expListBook As Workbook
Public lastAccRow As Integer
Public rowModify As Integer

Private Sub cmdAddNew_Click()
    rowModify = 0
    'frmNewAcc.Show
    expListBook.Activate
    Application.Run ("'" & expListBook.Name & "'!mTasks.ShowNewAccForm")
    Call PopAccLsb
End Sub

Private Sub cmdMod_Click()
    If lsbCurrAccs.ListIndex = -1 Then
        MsgBox "No item selected."
        Exit Sub
    End If

    accNum = lsbCurrAccs.ListIndex
    rowModify = 0
    
    For t = 0 To lsbCurrAccs.ListCount - 1
        If lsbCurrAccs.Selected(t) = True Then
            rowModify = t + 2
           blSelected = True
        End If
    Next t
    
    If blSelected = False Then
        MsgBox "Please select account import settings to modify."
        Exit Sub
    End If
    
    frmNewAcc.Show
End Sub

Private Sub cmdRemove_Click()
    If lsbCurrAccs.ListIndex = -1 Then
        MsgBox "No item selected"
        Exit Sub
    End If
    
    remYN = MsgBox("Are you sure you want to remove the selected settings?", vbYesNo)
    
    If remYN = vbYes Then
        For t = 0 To lsbCurrAccs.ListCount - 1
            If lsbCurrAccs.Selected(t) = True Then
                accVarWs.Cells(t + 2, 1).EntireRow.Delete
            End If
        Next t
    
        Call PopAccLsb
    End If
    
End Sub

Private Sub UserForm_Activate()
    Call PopAccLsb
End Sub

Private Sub UserForm_Initialize()
    Set accVarWs = ActiveWorkbook.Sheets("Account Variables")
    Set expListBook = ActiveWorkbook
    Call PopAccLsb
End Sub

Public Sub PopAccLsb()
    lsbCurrAccs.Clear
    lastAccRow = GetEmptyRow(accVarWs, 1, 2) - 1
    c = 0
    
    For accRow = 2 To lastAccRow
        lsbCurrAccs.AddItem
        lsbCurrAccs.List(c, 0) = accVarWs.Cells(accRow, 1).Value
        lsbCurrAccs.List(c, 1) = accVarWs.Cells(accRow, 2).Value
        c = c + 1
    Next accRow
End Sub


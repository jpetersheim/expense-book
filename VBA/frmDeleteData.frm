VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDeleteData 
   Caption         =   "Select data to delete"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5985
   OleObjectBlob   =   "frmDeleteData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDeleteData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public datesArr As Variant
Public filesArr As Variant
Public accountsArr As Variant
Public condCol As Integer

Private Sub cmbDDGroup_Change()
    Call ReloadDatasetListBox
End Sub

Private Sub cmbDDDelete_Change()
    cmbDDGroup.Clear
    
    If cmbDDDelete.Value = "Added Date" Then
        cmbDDGroup.List = datesArr
    ElseIf cmbDDDelete.Value = "Account" Then
        cmbDDGroup.List = accountsArr
    End If
    
    cmbDDGroup.Enabled = True
End Sub

Private Sub cmbSetType_Change()
    Call ReloadDatasetListBox
End Sub

Private Sub cmdCancel_Click()
    expListBook.Sheets("Main Tab").Activate
    Unload frmDeleteData
End Sub

Private Sub cmdDelete_Click()

    If mpDelete.Value = 0 Then
        Set lsbData = lsbDataset
    Else
        Set lsbData = lsbDDDataset
    End If

    If lsbData.ListIndex = -1 Then
        MsgBox "No item selected."
        Exit Sub
    End If
    
    deleteYN = MsgBox("Are you sure you want to delete the selected files?", vbYesNo)
    
    If deleteYN = vbNo Then
        Exit Sub
    End If
    
    firstRow = 3
    lastRow = expListSheet.Range("A65536").End(xlUp).Row
    
    If mpDelete.Value = 0 Then
        Set lsbData = lsbDataset

        Select Case cmbSetType.Value
            Case "Added Date"
                checkCol = 11
            Case "Account"
                checkCol = 10
            Case "Source File"
                checkCol = 12
        End Select
        
        For i = 0 To lsbData.ListCount - 1
            If lsbData.Selected(i) Then
                For j = firstRow To lastRow
                    'Can format - strings still show up as their original values, just dates are changed
                    If Format(expListSheet.Cells(j, checkCol).Value, "dd-mmm-yyyy") = lsbData.List(i) Then
                        expListSheet.Rows(j).Delete
                        j = j - 1
                    End If
                Next j
            End If
        Next i
    
    Else
        Set lsbData = lsbDDDataset
        
        Select Case cmbDDDelete.Value
            Case "Added Date"
                checkCol = 11
            Case "Account"
                checkCol = 10
        End Select
        
        check2 = cmbDDGroup.Value
        
        For i = 0 To lsbData.ListCount - 1
            If lsbData.Selected(i) Then
                For j = firstRow To lastRow
                    'Can format - strings still show up as their original values, just dates are changed
                    If Format(expListSheet.Cells(j, 12).Value, "dd-mmm-yyyy") = lsbData.List(i) And _
                    Format(expListSheet.Cells(j, checkCol).Value, "dd-mmm-yyyy") = check2 Then
                        expListSheet.Rows(j).Delete
                        j = j - 1
                    End If
                Next j
            End If
        Next i
        
        Call cmbDDDelete_Change
        
    End If
    
    Call SetTransIDs
    Call ReloadDatasetListBox
    
End Sub

Private Sub UserForm_Initialize()

    cmbSetType.Clear
    cmbSetType.AddItem "Added Date"
    cmbSetType.AddItem "Account"
    cmbSetType.AddItem "Source File"
    
    cmbSetType.Value = "Added Date"
    
    cmbDDGroup.Clear
    cmbDDDelete.Clear
    cmbDDGroup.Enabled = False
    cmbDDDelete.AddItem "Added Date"
    cmbDDDelete.AddItem "Account"
    
    cmbDDDelete.Value = "Added Date"
    
    
End Sub

Private Sub ReloadDatasetListBox()

    firstRow = 3
    lastRow = expListSheet.Range("A65536").End(xlUp).Row
    
    If lastRow < 3 Then
        lastRow = 3
    End If
    
    datesArr = GetUnique(expListSheet.Range(Cells(firstRow, 11), Cells(lastRow, 11)))
    accountsArr = GetUnique(expListSheet.Range(Cells(firstRow, 10), Cells(lastRow, 10)))
    filesArr = GetUnique(expListSheet.Range(Cells(firstRow, 12), Cells(lastRow, 12)))
    
    For D = 0 To ArrayLen(datesArr)
        datesArr(D) = Format(datesArr(D), "dd-mmm-yyyy")
    Next D
    
    lsbDataset.Clear
    
    If mpDelete.Value = 0 Then
        Set lsbData = lsbDataset
        Set cmbData = cmbSetType
        
        If cmbData.Value = "Added Date" Then
            tempArr = datesArr
        ElseIf cmbData.Value = "Account" Then
            tempArr = accountsArr
        ElseIf cmbData.Value = "Source File" Then
            tempArr = filesArr
        End If
        
    Else
        Set lsbData = lsbDDDataset
        Set cmbData = cmbDDGroup
        
        If cmbDDDelete.Value = "Added Date" Then
            condCol = 11
            rngCol = 12
            If cmbData.Value <> "" Then
                condFind = Format(cmbData.Value, "General Number")
                condFind = CLng(condFind)
            End If
        ElseIf cmbDDDelete.Value = "Account" Then
            condCol = 10
            rngCol = 12
            condFind = cmbData.Value
        Else
            condCol = 44
            rngCol = 44
        End If
        
        tempArr = GetUniqueIf(expListSheet.Range(Cells(firstRow, rngCol), Cells(lastRow, rngCol)), _
            expListSheet.Range(Cells(firstRow, condCol), Cells(lastRow, condCol)), condFind)
        
    End If

    lsbData.Clear
    lsbData.List = tempArr

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    expListBook.Sheets("Main Tab").Activate
    Unload frmDeleteData
End Sub

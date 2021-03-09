Attribute VB_Name = "mTasks"
Public Sub GetCategs()
    Dim categArr As Variant
    Dim currCategArr As Variant
    Dim addCateg As Boolean
    
    Application.ScreenUpdating = False

    Set expListBook = ActiveWorkbook
    Set expListSheet = expListBook.Sheets("Expense List")
    
    expListSheet.Activate
    
    LastRow = expListSheet.Range("A65536").End(xlUp).Row
    categArr = GetUnique(expListSheet.Range(Cells(3, 6), Cells(LastRow, 6)))
    
    expListBook.Sheets("Main Tab").Activate
    
    For j = LBound(categArr) To UBound(categArr)
        
        emptyCategRow = GetEmptyRow(expListBook.Sheets("Main Tab"), 6, 11)
        currCategArr = GetUnique(expListBook.Sheets("Main Tab").Range(Cells(11, 6), Cells(emptyCategRow - 1, 6)))
        addCateg = True
    
        For k = LBound(currCategArr) To UBound(currCategArr)
            If categArr(j) = currCategArr(k) Then
                addCateg = False
            End If
        Next k
        
        If addCateg = True Then
            expListBook.Sheets("Main Tab").Cells(emptyCategRow, 6) = categArr(j)
        End If
        
    Next j
    
    arrCats = GetUnique(Sheets("Main Tab").Range("F11", Range("F11").End(xlDown)))
    
    Sheets("Working Sheet").Activate
    Sheets("Working Sheet").Range("D5", Range("D5").End(xlDown)).Clear
    Sheets("Working Sheet").Cells(5, 4).Resize(UBound(arrCats) + 1) = WorksheetFunction.Transpose(arrCats)
    
    LastRow = GetEmptyRow(Sheets("Working Sheet"), 4, 3) - 1
    ActiveWorkbook.Names("Cat_List").RefersTo = Sheets("Working Sheet").Range(Cells(3, 4), Cells(LastRow, 4))
    
    Sheets("Main Tab").Activate
    Application.ScreenUpdating = True
    
End Sub

Attribute VB_Name = "Delete_Duplicates"
Sub Delete_Duplicates()
    
    Application.ScreenUpdating = False
    
    Dim RowN As Integer
    Dim RowM As Integer
    Dim ColN As Integer
    Dim LastRow As Integer, LastDataRow As Integer
    
    Dim Tbl As ListObject
    
    Set Tbl = ActiveSheet.ListObjects(1)
    
    ColN = Tbl.ListColumns("Duplicate Check").Index
    LastRow = Cells(Rows.Count, "C").End(xlUp).Row
    LastDataRow = LastRow
     
    For RowN = 2 To LastRow
        RowM = RowN - 1
        If Cells(RowN, ColN).Value = Cells(RowM, ColN).Value Then
            Rows(RowN).Delete
            RowN = RowN - 1
            LastDataRow = LastDataRow - 1
        End If
''        Debug.Print RowN

        If RowN > LastDataRow Then
            Exit For
        End If
    Next
    
   Application.ScreenUpdating = True

End Sub



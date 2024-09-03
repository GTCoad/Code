Public Sub DeleteStyles()

Dim deleted As Integer
Dim timeStart, timeFinish As Double



RestartTime:
timeStart = Timer
Restart:

If deleted = 200 Then
    
    timeFinish = Timer
    
    MSG1 = MsgBox(deleted & " Styles deleted in " & Format(timeFinish - timeStart, "0.00 \s\ec") & " seconds, " & ActiveWorkbook.Styles.Count & " styles left, Continue?", vbYesNo, "100 Deleted")
    If MSG1 = vbYes Then
      deleted = 0
      GoTo Restart
    Else
      Exit Sub
    End If
    
End If

    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "2") > 0 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "3") > 0 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "Normal ") > 0 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "F") = 1 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "Style") = 1 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "?") = 1 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "Audit") = 1 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
        For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "Dollar") = 1 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "ead") > 0 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "%") > 0 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "Co") > 0 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "Cur") > 0 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "Acc") > 0 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "Per") > 0 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
    For i = 1 To ActiveWorkbook.Styles.Count
        If InStr(ActiveWorkbook.Styles(i).Name, "-") > 0 Then
            ActiveWorkbook.Styles(i).Delete
            deleted = deleted + 1
            GoTo Restart
        End If
    Next
End Sub
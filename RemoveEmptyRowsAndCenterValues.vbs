Sub RemoveEmptyRowsAndCenterValues()
    Dim lastRow As Long
    Dim i As Long
    
    ' Находим последнюю строку в таблице
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Удаляем строки, если значения в столбцах B, C и D пустые
    For i = lastRow To 2 Step -1
        If Cells(i, "B").Value = "" And Cells(i, "C").Value = "" And Cells(i, "D").Value = "" Then
            Rows(i).Delete
        End If
    Next i
    
    ' Выравниваем все ячейки таблицы по центру и по горизонтали и по вертикали
    Range("A1:E" & lastRow).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub RemoveDuplicatesAndBlanks()
    Dim lastRow As Long
    Dim rng As Range
    Dim i As Long
    
    'Устанавливаем диапазон для работы
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row
    Set rng = Range("B2:D" & lastRow)
    
    'Сортируем по столбцам B, C и D
    rng.Sort key1:=Range("B2"), order1:=xlAscending, _
              key2:=Range("C2"), order2:=xlAscending, _
              key3:=Range("D2"), order3:=xlAscending, _
              Header:=xlYes
    
    'Удаляем дубликаты в столбцах B, C и D
    rng.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
          
    'Проверяем, не являются ли ячейки B3, C3 и D3 дубликатами B2, C2 и D2
    If Range("B2").Value = Range("B3").Value And _
       Range("C2").Value = Range("C3").Value And _
       Range("D2").Value = Range("D3").Value Then
        Range("B3:D3").ClearContents
    End If

    'Удаляем пустые строки
    For i = lastRow To 2 Step -1
        If WorksheetFunction.CountA(rng.Rows(i - 1)) = 0 Then
            rng.Rows(i - 1).EntireRow.Delete
        End If
    Next i	

End Sub

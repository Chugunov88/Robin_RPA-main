Sub SortAndNumberAndSum()
    Dim lastRow As Long
    Dim i As Long
    Dim currentValue As String
    Dim total As Double
    Dim mergeRange As Range
    Dim deleteRange As Range
    ' Находим последнюю строку в таблице
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Сортируем диапазон данных по столбцу B, затем по столбцу C, затем по столбцу D
    Range("A1:E" & lastRow).Sort key1:=Range("B2"), order1:=xlAscending, key2:=Range("C2"), order2:=xlAscending, key3:=Range("D2"), order3:=xlAscending, Header:=xlYes

    ' Проходимся по каждой строке и считаем сумму значений в столбце E для одинаковых значений в столбцах B, C и D
    currentValue = Cells(2, "B").Value & Cells(2, "C").Value & Cells(2, "D").Value
    total = Cells(2, "E").Value
    Set mergeRange = Cells(2, "E")

    For i = 3 To lastRow
        If Cells(i, "B").Value & Cells(i, "C").Value & Cells(i, "D").Value = currentValue Then
            total = total + Cells(i, "E").Value
            Cells(i, "E").Value = Cells(i, "F").Value ' Копируем значение из столбца F в столбец E
            Set mergeRange = Union(mergeRange, Cells(i, "E"))
        Else
            ' Объединяем ячейки со значениями и выравниваем содержимое по центру
            mergeRange.Merge
            mergeRange.HorizontalAlignment = xlCenter

            ' Записываем номер строки, сумму и значение в столбцы A, E соответственно
            Cells(i - 1, "A").Value = i - 2
            Cells(i - 1, "E").Value = total
            

            ' Обновляем значения для текущей группы
            currentValue = Cells(i, "B").Value & Cells(i, "C").Value & Cells(i, "D").Value
            total = Cells(i, "E").Value
            Set mergeRange = Cells(i, "E")
        End If
    Next i

    ' Записываем последнюю группу
    mergeRange.Merge
    mergeRange.HorizontalAlignment = xlCenter
    Cells(lastRow - 1, "A").Value = lastRow - 2
    Cells(lastRow - 1, "E").Value = total
    

    ' Пронумеруем строки начиная со второй строки после сортировки
    For i = 2 To lastRow - 1
        Cells(i, "A").Value = i - 1
    Next i
End Sub

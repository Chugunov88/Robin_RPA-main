Sub macros()
    Dim lastRow As Long
    Dim i As Long
    Dim currentValue As String
    Dim total As Double
    Dim deleteRange As Range
    Dim rowIndex As Long
    
    ' Находим последнюю строку в таблице
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Сортируем диапазон данных по столбцу B, затем по столбцу C, затем по столбцу D
    Range("A1:E" & lastRow).Sort key1:=Range("B2"), order1:=xlAscending, key2:=Range("C2"), order2:=xlAscending, key3:=Range("D2"), order3:=xlAscending, Header:=xlYes

    ' Проходимся по каждой строке и считаем сумму значений в столбце E для одинаковых значений в столбцах B, C и D
    currentValue = Cells(2, "B").Value & Cells(2, "C").Value & Cells(2, "D").Value
    total = Cells(2, "E").Value
    rowIndex = 2 ' Используется для запоминания строки последней обработанной группы

    For i = 3 To lastRow
        If Cells(i, "B").Value & Cells(i, "C").Value & Cells(i, "D").Value = currentValue Then
            total = total + Cells(i, "E").Value
            ' Удаляем строку если значения в столбцах B, C и D совпадают
            If deleteRange Is Nothing Then
                Set deleteRange = Cells(i, "A").EntireRow
            Else
                Set deleteRange = Union(deleteRange, Cells(i, "A").EntireRow)
            End If
        Else
            ' Записываем номер строки, сумму и значение в столбцы A, E соответственно
            Cells(rowIndex, "A").Value = rowIndex - 1
            Cells(rowIndex, "E").Value = total

            ' Обновляем значения для текущей группы
            currentValue = Cells(i, "B").Value & Cells(i, "C").Value & Cells(i, "D").Value
            total = Cells(i, "E").Value
            rowIndex = i
            
            ' Объединяем ячейки со значениями и выравниваем содержимое по центру
            Range("E" & rowIndex - 1).Value = total
            Range("E" & rowIndex - 1).HorizontalAlignment = xlCenter
        End If
    Next i
    
    ' Удаляем последнюю группу
    If Not deleteRange Is Nothing Then
        deleteRange.Delete Shift:=xlUp
    End If

    ' Пронумеруем строки начиная со второй строки после сортировки
    For i = 2 To lastRow - 1
        Cells(i, "A").Value = i - 1
    Next i
End Sub

Sub SortIfDigitMatches()
    Dim lastRow As Long
    Dim i As Long
    Dim cellB As String, cellC As String, cellD As String
    Dim sortRange As Range
    
    lastRow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' Проходимся по каждой строке и проверяем значения ячеек B, C и D
    For i = 2 To lastRow
        cellB = Right(ActiveSheet.Cells(i, "B"), 1)
        cellC = Right(ActiveSheet.Cells(i, "C"), 1)
        cellD = Right(ActiveSheet.Cells(i, "D"), 1)
        
        ' Если цифры совпадают, добавляем строку в диапазон сортировки
        If cellB = cellD And cellC = cellD Then
            If sortRange Is Nothing Then
                Set sortRange = Range("B" & i & ":C" & i)
            Else
                Set sortRange = Union(sortRange, Range("B" & i & ":C" & i))
            End If
        End If
    Next i
    
    ' Если есть строки для сортировки, выполняем сортировку
    If Not sortRange Is Nothing And sortRange.Rows.Count > 1 Then
        sortRange.Sort key1:=Range("B2"), header:=xlYes, MatchCase:=False
    End If
End Sub

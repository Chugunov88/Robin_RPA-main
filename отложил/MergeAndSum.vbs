Sub MergeAndSum()
    Dim i As Integer
    For i = 1 To Range("B" & Rows.Count).End(xlUp).Row - 1
        If Range("B" & i).Value = Range("B" & i + 1).Value And Range("C" & i).Value = Range("C" & i + 1).Value And Range("D" & i).Value = Range("D" & i + 1).Value Then
            Range("E" & i).Value = Range("E" & i).Value + Range("E" & i + 1).Value
            Range("E" & i + 1).ClearContents
            Range("E" & i & ":E" & i + 1).Merge
	    Range("B" & i & ":E" & i).Delete shift:=xlUp
        End If
    Next i
End Sub
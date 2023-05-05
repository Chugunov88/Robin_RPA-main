Sub macros()

Dim lRow As Long, lRow2 As Long, lRow3 As Long
Dim sName As String, sParam As String, sOrg As String
Dim sSum As String

Application.ScreenUpdating = False

lRow = Cells(Rows.Count, 1).End(xlUp).Row

For lRow2 = lRow To 2 Step -1
    sName = Cells(lRow2, 2).Value
    sParam = Cells(lRow2, 3).Value
    sOrg = Cells(lRow2, 4).Value
    sSum = Cells(lRow2, 5).Value
    
    For lRow3 = lRow2 - 1 To 2 Step -1
        If sName = Cells(lRow3, 2).Value And sParam = Cells(lRow3, 3).Value And sOrg = Cells(lRow3, 4).Value Then
            Cells(lRow3, 5).Value = Cells(lRow3, 5).Value + sSum
            Rows(lRow2).Delete
            Exit For
        End If
    Next lRow3
Next lRow2

Application.ScreenUpdating = True
Selection.HorizontalAlignment = xlCenter
Selection.VerticalAlignment = xlCenter

End Sub
Sub HW2()

Dim lRow, volume, counter As Integer
Dim open_price, close_price, yearly_change As Double
lRow = Cells(Rows.Count, "A").End(xlUp).Row
volume = 0
counter = 2
open_price = Cells(2, 3).Value
For i = 2 To lRow

    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        volume = volume + Cells(i, 7).Value
    ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        volume = volume + Cells(i, 7).Value
        Cells(counter, 9).Value = Cells(i, 1).Value
        Cells(counter, 10).Value = volume
        volume = 0
        counter = counter + 1
        close_price = Cells(i, 6).Value
        yearly_change = close_price - open_price
        Cells(counter - 1, 11).Value = yearly_change
        open_price = Cells(i + 1, 3).Value
    End If
Next i

End Sub

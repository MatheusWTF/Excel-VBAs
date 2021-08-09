Public Function Arrumar()
    Dim i As Integer, j As Integer, NumRows As Integer, c As Integer
    
    NumRows = Range("A1").End(xlDown).Row
    c = 1
    For i = 1 To NumRows Step 3
        
        For j = 1 To 3 Step 1
            Cells(((i + j) - 1), 1).Cut Cells(c, j)
        Next j
        c = c + 1
    Next i
    
End Function

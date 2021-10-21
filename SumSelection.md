Sub SumArray()

Dim i As Integer, j As Integer

Dim X As Variant, s As Double

X = Selection

For i = 1 To 4

    For j = 1 To 3

        s = s + X(j, i)

    Next j

Next i

MsgBox s

End Sub

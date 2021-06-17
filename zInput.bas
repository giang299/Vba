Attribute VB_Name = "zInput"
Function array_test_nxm(n As Integer, m As Integer, space As Integer)
Dim Arraytest() As Double
space = space - 1
ReDim Arraytest(1 To n, 1 To m)
For i = 1 To n - 1
For j = 1 To m
 Arraytest(i, j) = j + space
Next j
Arraytest(i, m) = 0
Next i
array_test_nxm = Arraytest()

End Function

Sub testfuntion()

E = zInput.array_test_nxm(3, 10, 2)

End Sub

8GWVIuv89LnNWR0R7+y8Ou7r7qOlO0BjcCmhRb72XULvvIENXgBX3HJtTtsY3sYYf1wNbOh55G9NSWLdCxcT3EJdViV54Uy+nIiq27eLfsHCoc
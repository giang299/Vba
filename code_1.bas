Attribute VB_Name = "code_1"
Sub FindLastrows_of_sheet4()
lr = s5.Cells(s5.Rows.Count, "A").End(xlUp).Row
With s5.Range("e1:G" & lr)
    .NumberFormat = "General"
    .Value = .Value
End With
For i = lr To lr + 1000
If Left(s4.Range("B" & i), 1) = "~" Then
LastRows4 = s4.Range("B" & i).Row - 1
Exit Sub
End If
Next i
End Sub

Sub Get_coordinate()
lr = s5.Cells(s3.Rows.Count, "A").End(xlUp).Row
ReDim point1(1 To 2, 2 To lr) As Double
ReDim point2(1 To 2, 2 To lr) As Double
For i = 1 To LastRows4
For j = 1 To LastRows4
If s5.Range("f" & i + 1) = s4.Range("e" & j + 3) Then
point1(1, i + 1) = s4.Range("f" & j + 3)
point1(2, i + 1) = s4.Range("g" & j + 3)
Exit For
End If
Next j
Next i
For i = 1 To LastRows4
For j = 1 To LastRows4
If s5.Range("g" & i + 1) = s4.Range("e" & j + 3) Then
point2(1, i + 1) = s4.Range("f" & j + 3)
point2(2, i + 1) = s4.Range("g" & j + 3)
Exit For
End If
Next j
Next i


End Sub

Sub change_to_array()
akhaibao.kbsheet

lr = s5.Cells(s3.Rows.Count, "A").End(xlUp).Row
Mmax = s5.Range("b2:c" & lr)
section = s5.Range("d2:d" & lr)
beam_label = s5.Range("a2:a" & lr)
End Sub
4  (   �O  �        �O  �        �O  �        �O  �        �O� �        �ON �        �O�	 �        �O�
 �        �O �        �O6 �        �O6 �        �O� �        �O� �        �O� �        �
O� �        �	O5 �        �O� �        �Oe �        �O� �        �O �        �O{ �5  (   �Of �        �O( �        �Oz �
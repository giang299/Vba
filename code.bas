Attribute VB_Name = "code"
Sub Momentmax()
lr = s3.Cells(s3.Rows.Count, "A").End(xlUp).Row
s5.Range("a2") = s3.Range("b4")
s5.Range("b2") = Format(s3.Range("o4"), "0.00")
stt = 2
For i = 1 To lr
'Beam name
If s3.Range("b" & i + 3).Value = s3.Range("b" & i + 4).Value Then
'Moment max
If s5.Range("b" & stt) < s3.Range("o" & i + 4) Then
s5.Range("b" & stt) = Format(s3.Range("o" & i + 4), "0.00")
End If
Else
stt = stt + 1
s5.Range("a" & stt) = s3.Range("b" & i + 4)
s5.Range("b" & stt) = Format(s3.Range("o" & i + 4), "0.00")
End If
Next i
End Sub

Sub Momentmin()
lr = s3.Cells(s3.Rows.Count, "A").End(xlUp).Row

Dim stt As Integer
stt = 2
s5.Range("c2") = Format(s3.Range("o4"), "0.00")
For i = 1 To lr
'beam name
If s3.Range("b" & i + 3) = s3.Range("b" & i + 4) Then
'moment min
If s5.Range("c" & stt) > s3.Range("o" & i + 4) Then
s5.Range("c" & stt) = Format(s3.Range("o" & i + 4), "0.00")
End If
Else
stt = stt + 1
s5.Range("c" & stt) = Format(s3.Range("o" & i + 4), "0.00")
End If

Next i
End Sub

 Sub Get_uniquenameAndSection_prop_from_label_beam()
lr = s2.Cells(Sheet1.Rows.Count, "A").End(xlUp).Row
With s2.Range("a1:G" & lr)
    .NumberFormat = "General"
    .Value = .Value
End With

For i = 1 To lr
If s2.Range("B" & i + 3) = s5.Range("a" & i + 1) Then
s5.Range("d" & i + 1) = s2.Range("f" & i + 3)
s5.Range("e" & i + 1) = s2.Range("c" & i + 3)

End If
Next i
End Sub

 Sub unique_beam_change_to2point()
 With s1.Range("a1:G" & lr)
    .NumberFormat = "General"
    .Value = .Value
End With
 
lr = s1.Cells(s1.Rows.Count, "A").End(xlUp).Row
For i = 1 To lr
If s5.Range("e" & i + 1) = s1.Range("a" & i + 3) Then
s5.Range("f" & i + 1) = s1.Range("e" & i + 3)
s5.Range("g" & i + 1) = s1.Range("f" & i + 3)
End If
Next i
End Sub




]O  ?        (Z??                          ?j^O  ?0&Z??  ??Y??  ?????                  ?j[O  ??'Z??  `&Z??  @%Z??               ?jTO  ??&Z??  P'Z??                ?        ?jQO  ?P*Z??  (Z??               @        ?jRO  ??'Z??  0)Z??  ?'Z??               ?jOO  ? )Z??  ?'Z??  ?)Z??               ?jHO  ?        `)Z??                          ?jEO  ??'Z??  0)Z??  ?)Z??               ?jFO  ??'Z??  ?)Z??  ?'Z??                ?jCO  ??????  ??Y??  ?????    
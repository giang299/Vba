Attribute VB_Name = "code_cad"
Sub GetObject_Acad()
    On Error Resume Next
    Set Acad = GetObject(, "AutoCAD.Application")
    If (Err <> 0) Then
        Err.Clear
        Set Acad = New AcadApplication
        If (Err <> 0) Then
            MsgBox Err.Number & " " & Err.Description
            End
        End If
    End If
    Acad.Visible = True
End Sub
Sub drawing_beam_max()
stt = 2
Dim p1x As Double
Dim p1y As Double
Dim p2x As Double
Dim p2y As Double
Dim M_max As AcadText
Dim newText As AcadText
Dim new_Line As AcadLine
Dim mid_point(0 To 2) As Double
For Each Item In beam_label
p1x = point1(1, stt) + rp_point(0)
p1y = point1(2, stt) + rp_point(1)
p2x = point2(1, stt) + rp_point(0)
p2y = point2(2, stt) + rp_point(1)
mid_point(0) = (p1x + p2x) / 2
mid_point(1) = (p1y + p2y) / 2
Set new_Line = func_cad.addline_2ponit_4coordinate(p1x, p1y, p2x, p2y)
Set newText = Acad.ActiveDocument.ModelSpace.AddText(Item, _
mid_point, 300)
mid_point(0) = mid_point(0) - 500
mid_point(1) = mid_point(1) - 500
Set M_max = Acad.ActiveDocument.ModelSpace.AddText(s5.Range("b" & stt), _
mid_point, 300)
stt = stt + 1
Next Item

End Sub
Sub drawing_beam_min()
stt = 2
Dim p1x As Double
Dim p1y As Double
Dim p2x As Double
Dim p2y As Double
Dim M_max As AcadText
Dim newText As AcadText
Dim new_Line As AcadLine
Dim mid_point(0 To 2) As Double
For Each Item In beam_label
p1x = point1(1, stt) + rp_point(0)
p1y = point1(2, stt) + rp_point(1)
p2x = point2(1, stt) + rp_point(0)
p2y = point2(2, stt) + rp_point(1)
mid_point(0) = (p1x + p2x) / 2
mid_point(1) = (p1y + p2y) / 2
Set new_Line = func_cad.addline_2ponit_4coordinate(p1x, p1y, p2x, p2y)
Set newText = Acad.ActiveDocument.ModelSpace.AddText(Item, _
mid_point, 300)
mid_point(0) = mid_point(0) - 500
mid_point(1) = mid_point(1) - 500
Set M_max = Acad.ActiveDocument.ModelSpace.AddText(s5.Range("c" & stt), _
mid_point, 300)
stt = stt + 1
Next Item

End Sub
Sub change_to_array()
akhaibao.kbsheet

lr = s5.Cells(s3.Rows.Count, "A").End(xlUp).Row
Mmax = s5.Range("b2:c" & lr)
section = s5.Range("d2:d" & lr)
beam_label = s5.Range("a2:a" & lr)
End Sub
          ����  zC  zC                                                 @�> @�>           :           :                  �?                                ����  zC  zC                                                 @�> @�>           :           :                  �?                                ����  zC  zC                                                 �? @�>           :
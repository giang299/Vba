Attribute VB_Name = "func_cad"
Function addline_2ponit_4coordinate(a As Double, b As Double, c As Double, d As Double) As AcadLine
 Dim PointStart(1 To 3) As Double
 Dim PointEnd(1 To 3) As Double
 PointStart(1) = a
 PointStart(2) = b
 PointEnd(1) = c
 PointEnd(2) = d
 Set addline_2ponit_4coordinate = Acad.ActiveDocument.ModelSpace.AddLine(PointStart, PointEnd)
    
End Function

Function Array_to_object(A_Array() As Double, aucad As AcadApplication)
Dim stt As Integer
Dim point1(1 To 3) As Double
Dim point2(1 To 3) As Double
Dim d As AcadLine
stt = 1
Dim lr1 As Integer
lr1 = UBound(A_Array, 1)
For i = 1 To lr1
point1(1) = A_Array(1, i)
point1(2) = A_Array(2, i)
point1(1) = A_Array(3, i)
point1(2) = A_Array(4, i)
Set d = Acad.ActiveDocument.ModelSpace.AddLine(point1, point2)
Next i
End Function




Function addtext_ToCoordinate(a As Double, b As Double) As AcadText
 Dim Point(1 To 3) As Double
 Set addtext_ToCoordinate = Acad.ActiveDocument.ModelSpace. _
 AddText(Text, Point, 300)
End Function





	       	^/           Ó¶ÿù     øÿÿÿÿ              ?¥O           ¥O          ¥O          ¥O    (   ¥O          ¥O          i¥O          h¥O          k¥O          j¥O 	         m¥O 
         l¥O          o¥O          n¥O          a¥O          `¥ O          c¥O          b¥O   (   e¥OD         d¥Où         g¥O          f¥
O          y¥	O          x¥O          {¥O          z¥O          }¥O  
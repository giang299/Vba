Attribute VB_Name = "set_Layer"
Sub make_Layer()
Dim newLayer As AcadLayer
On Error Resume Next
    Set newLayer = Acad.ActiveDocument.Layers.Add("Beam")
    newLayer.Color = acWhite
    Set newLayer = Acad.ActiveDocument.Layers.Add("Column")
    newLayer.Color = acBlue
    Set newLayer = Acad.ActiveDocument.Layers.Add("Section")
    newLayer.Color = acBlue
    Set newLayer = Acad.ActiveDocument.Layers.Add("Value")
    newLayer.Color = acBlue
On Error GoTo 0
End Sub

Sub activate_Beam()
Acad.ActiveDocument.ActiveLayer = Acad.ActiveDocument.Layers.Add("Beam")
End Sub
Sub activate_Column()
Acad.ActiveDocument.ActiveLayer = Acad.ActiveDocument.Layers.Add("Column")
End Sub
Sub activate_Value()
Acad.ActiveDocument.ActiveLayer = Acad.ActiveDocument.Layers.Add("Value")
End Sub
Sub activate_Section()
Acad.ActiveDocument.ActiveLayer = Acad.ActiveDocument.Layers.Add("Section")
End Sub

 "0.00")
End If
Else
stt = stt + 1
s5.Range("c" & stt) = Format(s3.Range("o" & i + 4), "0.00")
End IftSourceL7  7^҃  j???  ??N??  rContentDependentState0  Data.DurationInMilliseconds0 ?? Data.InputType0  #Data.IsCreatedByIntelligentServices0 Data.IsSingleLineDisplay0 Data.IsSingleLineOverflow0 Data.ParentSurface0  Data.ParentTCID0 ??? Data.StartTime0???ӊ?Ԙ? Data.Success0? 	Data.TCID0 ?? Data.UserActionID0 ? Device.OsBuildi19042 Device.OsVersioni10.0 Event.ContractiOffice
Attribute VB_Name = "Main_code"


Public s1 As Worksheet
Public s2 As Worksheet
Public s3 As Worksheet
Public s4 As Worksheet
Public s5 As Worksheet
Public lr As Integer
Public stt As Integer
Public Acad As AcadApplication
Public point1 As Variant
Public point2 As Variant
Public LastRows4 As Integer
Public beam_label As Variant
Public Moment As Variant
Public section As Variant
Public rp_point As Variant
Sub CodeMain()
If sheetExists("beam-prop-ultimate") = False Then
Sheets.Add.Name = "beam-prop-ultimate"
End If
akhaibao.kbsheet
s5.Range("a2:I" & 1048576).ClearContents
s5.Range("a2:I" & 1048576).NumberFormat = "General"
'code
code.Momentmax
code.Momentmin
code.Get_uniquenameAndSection_prop_from_label_beam
code.unique_beam_change_to2point
'code_1
code_1.FindLastrows_of_sheet4
code_1.Get_coordinate
code_1.change_to_array
code_cad.GetObject_Acad
set_Layer.make_Layer
rp_point = Event_cad.get_point
set_Layer.activate_Beam
code_cad.drawing_beam_max
rp_point = Event_cad.get_point
code_cad.drawing_beam_min
Acad.ZoomExtents

On Error Resume Next
s5.Select
If Err <> 0 Then
MsgBox "khong_du_sheet_ok"
Exit Sub
End If
On Error GoTo 0
End Sub









                                                                                                                                                                                                                                                                               :   0 ,   ??o ?P,??  p?'??  G "   :   " O p 
Attribute VB_Name = "tcad"

Sub new_Layer()

Dim acadApp As AcadApplication
Dim acadDoc As AcadDocument
Dim acadDoc1 As AcadDocument

Dim newLayer As AcadLayer
Dim layerName As String

On Error Resume Next
    layerName = "Layer_1"
    Set newLayer = acadDoc.Layers.Add(layerName)
    newLayer.Color = acBlue
On Error GoTo 0


If newLayer Is Nothing Then
    Set newLayer = acadDoc1.Layers.Add(layerName)
End If
acadDoc1.ActiveLayer = newLayer
End Sub





                           
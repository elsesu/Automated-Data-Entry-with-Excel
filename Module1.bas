Attribute VB_Name = "Module1"
Private Sub Worksheet_Change(ByVal Target As Range)

Dim isect As Range
Set isect = Application.Intersect(Target, Range("H1:C20"))
If Not (isect Is Nothing) Then
    If Target.Value > 0 Then Target.Value = 0 - Target.Value
End If

End Sub



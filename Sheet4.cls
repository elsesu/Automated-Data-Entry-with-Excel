VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

Dim isect As Range
Set isect = Application.Intersect(Target, Range("K8:K197"))
If Not (isect Is Nothing) Then
    If Target.Value > 0 Then Target.Value = 0 - Target.Value
    
End If

End Sub

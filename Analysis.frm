VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Analysis 
   Caption         =   "Visualization"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600.001
   OleObjectBlob   =   "Analysis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbcatig_Change()
Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Input data") ' replace "Sheet" with the name of your worksheet
    Dim searchValue As String
    searchValue = cmbcatig.Value
    Dim foundRange As Range
    Set foundRange = ws.Range("A1:A100").Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundRange Is Nothing Then
        ws.Activate
        foundRange.Select
    End If
End Sub

Private Sub cmbmeall_Change()
Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Input data") ' replace "Sheet" with the name of your worksheet
    Dim searchValue As String
    searchValue = cmbmeall.Value
    Dim foundRange As Range
    Set foundRange = ws.Range("A1:A100").Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundRange Is Nothing Then
        ws.Activate
        foundRange.Select
    End If
End Sub

Private Sub EXIT5_Click()
Unload Me
End Sub


Private Sub txtEndDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtEndDate.Value) Then
        MsgBox "Please enter a valid date.", vbExclamation, "Invalid Input"
        Cancel = True
    End If

End Sub

 Private Sub txtStartDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtStartDate.Value) Then
        MsgBox "Please enter a valid date.", vbExclamation, "Invalid Input"
        Cancel = True
    End If
End Sub

Private Sub UserForm_Initialize()
    
    make
    
End Sub

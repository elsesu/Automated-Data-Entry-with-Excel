VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} interface 
   Caption         =   "MENU"
   ClientHeight    =   8640.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13710
   OleObjectBlob   =   "interface.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.Show
End Sub

Private Sub cmdTOANALYZE_Click()
Analysis.Show
End Sub

Private Sub cmdAsAdmin_Click()
Me.Hide
loginform.Show
End Sub

Private Sub cmdTODASHBOARDS_Click()
optionform.Show
End Sub

Private Sub cmdTOEXIT_Click()
 Unload Me
 Application.DisplayAlerts = False
 Application.Quit
End Sub

Private Sub cmdTOINPUT_Click()
FrmForm.Show
End Sub

Private Sub cmdTOSTOCKINPUT_Click()
INVFORM.Show
End Sub

Private Sub CommandButton7_Click()

End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} optionform 
   Caption         =   "Selection"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6660
   OleObjectBlob   =   "optionform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "optionform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGENERAL_Click()
Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard") 'Change "Sheet1" to the name of the sheet you want to open'
    ws.Activate
    Application.Visible = True ' make Excel visible on screen
    Me.Hide
End Sub

Private Sub cmdinventory_Click()
Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("inventory dashboard") 'Change "Sheet1" to the name of the sheet you want to open'
    ws.Activate
    Application.Visible = True ' make Excel visible on screen
    Me.Hide
End Sub

Private Sub EXIT3_Click()
Unload Me
End Sub

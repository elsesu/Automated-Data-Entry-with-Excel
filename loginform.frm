VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} loginform 
   Caption         =   "UserForm2"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8145
   OleObjectBlob   =   "loginform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "loginform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtCLEAR_Click()
Me.txtUSERID.Value = ""
Me.txtUSERPASSWORD.Value = ""

Me.txtUSERID.SetFocus
End Sub

Private Sub txtLOGIN_Click()
    ' Get entered credentials from loginform
    Dim user As String
    Dim password As String
    user = Me.txtUSERID.Value
    password = Me.txtUSERPASSWORD.Value
    
    ' Check if credentials are valid
    If ValidateCredentials(user, password) Then
        ' Check if user is an admin
        If user = "admin" And password = "admin" Then
            ' User is admin - open adminform
            Unload Me
            AdminCenter.Show
        Else
            ' User is not admin - open userform
            Unload Me
            interface.Show
        End If
    Else
        ' Credentials are invalid - show error message
        MsgBox "Invalid username or password. Please try again.", vbCritical + vbOKOnly, "Error"
        Me.txtUSERPASSWORD.Value = ""
        Me.txtUSERPASSWORD.SetFocus
    End If
End Sub

Private Function ValidateCredentials(user As String, password As String) As Boolean
    Dim loginDB As Worksheet
    Set loginDB = ThisWorkbook.Sheets("Users")
    
    Dim userRow As Range
    Set userRow = loginDB.Columns("A").Find(user, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not userRow Is Nothing Then
        ' Check if password matches for the found user
        If loginDB.Cells(userRow.Row, "B").Value = password Then
            ValidateCredentials = True
            Exit Function
        End If
    End If
    
    ' Check if entered credentials match admin username and password
    If user = "admin" And password = "admin" Then
        ValidateCredentials = True
        Exit Function
    End If
    
    ' Credentials are invalid
    ValidateCredentials = False
End Function


Private Sub UserForm_Initialize()

Me.txtUSERID.Value = ""
Me.txtUSERPASSWORD.Value = ""

Me.txtUSERID.SetFocus

End Sub

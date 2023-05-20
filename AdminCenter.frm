VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminCenter 
   Caption         =   "adminform"
   ClientHeight    =   8325.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10650
   OleObjectBlob   =   "AdminCenter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdminCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnalysis_Click()
Analysis.Show
End Sub

Private Sub cmdbckToUser_Click()
 Me.Hide
 interface.Show
End Sub

Private Sub cmddashboards_Click()
optionform.Show
End Sub
Private Sub cmdDisplayUsers_Click()
    Dim ws As Worksheet
    Dim frm As Object ' declare variable for userform
    
    Set ws = ThisWorkbook.Sheets("Users") 'Change "Sheet1" to the name of the sheet you want to open'
    ws.Activate
    Application.Visible = True ' make Excel visible on screen
    Me.Hide
End Sub



Private Sub cmdEXIT6_Click()
Unload Me
End Sub

Private Sub cmdINPUT_Click()
FrmForm.Show
End Sub

Private Sub cmdinventory_Click()
INVFORM.Show
End Sub

Private Sub cmdUpload_Click()
ThisWorkbook.Save
End Sub

Private Sub UserForm_Initialize()
    LoadUsers
End Sub
Private Sub LoadUsers()
    ' Load the list of users from a worksheet named "Users"
    Dim lastRow As Long
    Dim userRange As Range
    Dim userCell As Range
    
    lastRow = Worksheets("Users").Cells(Rows.Count, "A").End(xlUp).Row
    Set userRange = Worksheets("Users").Range("A2:A" & lastRow)
    
    Me.LstUsers.Clear
    For Each userCell In userRange
        Me.LstUsers.AddItem userCell.Value
    Next userCell
End Sub
' Add new user to the login database
Private Sub AddUserToLoginDB(user As String, password As String)
    Dim loginDB As Worksheet
    Set loginDB = ThisWorkbook.Sheets("Users")
    
    ' Check if user already exists in the login database
    Dim existingUser As Range
    Set existingUser = loginDB.Columns("A").Find(user, LookIn:=xlValues, LookAt:=xlWhole)
    If Not existingUser Is Nothing Then
        MsgBox "Username already exists. Please choose a different username.", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
        If password <> Me.txtConfirmPassword.Value Then
        MsgBox "The passwords do not match. Please try again.", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If

    ' Add new user to login database
    Dim lastRow As Long
    lastRow = loginDB.Cells(Rows.Count, "A").End(xlUp).Row
    
    With loginDB
        .Cells(lastRow + 1, "A").Value = user
        .Cells(lastRow + 1, "B").Value = password
    End With
End Sub

Private Sub btnAddUser_Click()
' Get new user's credentials from adminform
    Dim user As String
    Dim password As String
    user = Me.txtNewUser.Value
    password = Me.txtNewPassword.Value
    
    ' Add new user to login database
    AddUserToLoginDB user, password
    
    ' Update list of users on adminform
    LoadUsers
End Sub

Private Sub btnDeleteUser_Click()
    ' Delete the selected user from the worksheet
    Dim selectedUser As String
    Dim userRow As Long
    
    If Me.LstUsers.ListIndex < 0 Then
        MsgBox "Please select a user to delete.", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    
    selectedUser = Me.LstUsers.Value
    userRow = WorksheetFunction.Match(selectedUser, ThisWorkbook.Sheets("Users").Range("A:A"), 0)
    
    If Not IsError(userRow) Then
        ThisWorkbook.Sheets("Users").Rows(userRow).Delete
        Me.LstUsers.RemoveItem Me.LstUsers.ListIndex
    Else
        MsgBox "Could not find user to delete.", vbCritical + vbOKOnly, "Error"
    End If
End Sub

    


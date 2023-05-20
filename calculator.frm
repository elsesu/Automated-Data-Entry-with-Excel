VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} calculator 
   Caption         =   "calculator"
   ClientHeight    =   8910.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   OleObjectBlob   =   "calculator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fstNum As Long
Dim SecNum As Long
Dim Oper As String
    
Private Sub CMB0_Click()
If (TXTDISPLAY.Text = "0") Then
    TXTDISPLAY.Text = "0"
Else
    TXTDISPLAY.Text = TXTDISPLAY.Text + "0"
End If
End Sub

Private Sub cmb1_Click()
If (TXTDISPLAY.Text = "0") Then
    TXTDISPLAY.Text = "1"
Else
    TXTDISPLAY.Text = TXTDISPLAY.Text + "1"
End If
End Sub

Private Sub cmb2_Click()
If (TXTDISPLAY.Text = "0") Then
    TXTDISPLAY.Text = "2"
Else
    TXTDISPLAY.Text = TXTDISPLAY.Text + "2"
End If
End Sub

Private Sub cmb3_Click()
If (TXTDISPLAY.Text = "0") Then
    TXTDISPLAY.Text = "3"
Else
    TXTDISPLAY.Text = TXTDISPLAY.Text + "3"
End If
End Sub

Private Sub cmb4_Click()
If (TXTDISPLAY.Text = "0") Then
    TXTDISPLAY.Text = "4"
Else
    TXTDISPLAY.Text = TXTDISPLAY.Text + "4"
End If
End Sub

Private Sub cmb5_Click()
If (TXTDISPLAY.Text = "0") Then
    TXTDISPLAY.Text = "5"
Else
    TXTDISPLAY.Text = TXTDISPLAY.Text + "5"
End If
End Sub

Private Sub cmb6_Click()
If (TXTDISPLAY.Text = "0") Then
    TXTDISPLAY.Text = "6"
Else
    TXTDISPLAY.Text = TXTDISPLAY.Text + "6"
End If
End Sub

Private Sub cmb7_Click()

If (TXTDISPLAY.Text = "0") Then
    TXTDISPLAY.Text = "7"
Else
    TXTDISPLAY.Text = TXTDISPLAY.Text + "7"
End If

End Sub

Private Sub cmb8_Click()

If (TXTDISPLAY.Text = "0") Then
    TXTDISPLAY.Text = "8"
Else
    TXTDISPLAY.Text = TXTDISPLAY.Text + "8"
End If


End Sub

Private Sub cmb9_Click()
If (TXTDISPLAY.Text = "0") Then
    TXTDISPLAY.Text = "9"
Else
    TXTDISPLAY.Text = TXTDISPLAY.Text + "9"
End If
End Sub

Private Sub CMBADD_Click()
fstNum = TXTDISPLAY.Text
Oper = "+"
TXTDISPLAY.Text = ""
End Sub

Private Sub CMBBackspace_Click()
TXTDISPLAY.Text = Left(TXTDISPLAY.Text, Len(TXTDISPLAY.Text) - 1)

End Sub

Private Sub cmbC_Click()
TXTDISPLAY.Text = "0"
End Sub

Private Sub CmbCE_Click()
Dim f, s As String

TXTDISPLAY.Text = "0"

f = fstNum
s = SecNum
f = ""
s = ""
End Sub

Private Sub CommandButton22_Click()

End Sub

Private Sub CMBDIV_Click()
fstNum = TXTDISPLAY.Text
Oper = "/"
TXTDISPLAY.Text = ""
End Sub

Private Sub cmbequals_Click()
 SecNum = TXTDISPLAY.Text
 
 Select Case Oper
 
 Case "+"
     TXTDISPLAY.Text = fstNum + SecNum
 Case "-"
     TXTDISPLAY.Text = fstNum - SecNum
 
 Case "*"
     TXTDISPLAY.Text = fstNum * SecNum
 Case "/"
 If SecNum = 0 Then
 TXTDISPLAY.Text = "Cannot divide by zero"
 Else
     TXTDISPLAY.Text = fstNum / SecNum
 End If
 End Select
End Sub

Private Sub CMBMINUS_Click()
fstNum = TXTDISPLAY.Text
Oper = "-"
TXTDISPLAY.Text = ""
End Sub

Private Sub CMBpoint_Click()
If InStr(TXTDISPLAY.Text, ".") = 0 Then
TXTDISPLAY.Text = TXTDISPLAY.Text + "."
End If
End Sub

Private Sub CMBTIMES_Click()
fstNum = TXTDISPLAY.Text
Oper = "*"
TXTDISPLAY.Text = ""
End Sub

Private Sub plusminus_Click()
Dim q
q = TXTDISPLAY.Text
TXTDISPLAY.Text = (-1 * q)

End Sub


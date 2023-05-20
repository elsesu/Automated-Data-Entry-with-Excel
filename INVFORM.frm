VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} INVFORM 
   Caption         =   "Inventory Form"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12870
   OleObjectBlob   =   "INVFORM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "INVFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox7_Change()

End Sub



Private Sub delete1_Click()
If selected_form = 0 Then
        MsgBox "No row is selected.", vbOKOnly + vbInformation, "Delete"
        Exit Sub
    End If
    
    Dim i As VbMsgBoxResult
    
    i = MsgBox("do you want to delete the selected records?", vbYesNo + vbQuestion, "Confirmation")
    
    If i = vbNo Then Exit Sub
    
    ThisWorkbook.Sheets("InventoryTesting").Rows(selected_form + 1).Delete
    
    Call Redo
    
    MsgBox "Selected record has been deleted", vbOKOnly + vbInformation, "Deleted"
End Sub

Private Sub edit1_Click()
If selected_form = 0 Then
         MsgBox "Now row is selected.", vbOKOnly + vbInformation, "Edit"
         Exit Sub
      End If
      'code to update the value to respective controls'
      
      Me.txtrowno.Value = selected_form + 1
      Me.cmbAdded.Value = Me.Invdatabase.List(Me.Invdatabase.ListIndex, 1)
      Me.cmbCosts.Value = Me.Invdatabase.List(Me.Invdatabase.ListIndex, 2)
      Me.cmbUsed.Value = Me.Invdatabase.List(Me.Invdatabase.ListIndex, 3)
      Me.cmbIngredient.Value = Me.Invdatabase.List(Me.Invdatabase.ListIndex, 4)
      Me.cmbCategory1.Value = Me.Invdatabase.List(Me.Invdatabase.ListIndex, 5)
            
      MsgBox "Please make required changes and save to update", vbOKOnly + vbInformation, "Edit"
End Sub

Private Sub Label9_Click()

End Sub

Private Sub EXIT2_Click()
Unload Me
End Sub

Private Sub txtAdd_Click()

Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Do you want to save the data?", vbYesNo + vbInformation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call Send
    Call Redo
End Sub

Private Sub txtReset1_Click()

   Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Do you want to reset the form?", vbYesNo + vbInformation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call Redo
End Sub

Private Sub UserForm_Initialize()
    
    Call Redo
    
End Sub

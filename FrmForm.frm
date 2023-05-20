VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmForm 
   Caption         =   "Input Data Form"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11325
   OleObjectBlob   =   "FrmForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub txtCATEGORY_Change()

End Sub

Private Sub cmdDELETE_Click()
    If selected_list = 0 Then
        MsgBox "No row is selected.", vbOKOnly + vbInformation, "Delete"
        Exit Sub
    End If
    
    Dim i As VbMsgBoxResult
    
    i = MsgBox("do you want to delete the selected records?", vbYesNo + vbQuestion, "Confirmation")
    
    If i = vbNo Then Exit Sub
    
    ThisWorkbook.Sheets("Database").Rows(selected_list + 1).Delete
    
    Call Reset
    
    MsgBox "Selected record has been deleted", vbOKOnly + vbInformation, "Deleted"
    
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub cmdEDIT_Click()
      If selected_list = 0 Then
         MsgBox "Now row is selected.", vbOKOnly + vbInformation, "Edit"
         Exit Sub
      End If
      'code to update the value to respective controls'
      Dim sGender As String
      
      Me.txtROWNUMBER.Value = selected_list + 1
      Me.txtID.Value = Me.bukadatabase.List(Me.bukadatabase.ListIndex, 1)
      Me.txtNAME.Value = Me.bukadatabase.List(Me.bukadatabase.ListIndex, 2)
      sGender = Me.bukadatabase.List(Me.bukadatabase.ListIndex, 3)
      
      If sGender = "Female" Then
         Me.optFEMALE.Value = True
      Else
         Me.optMALE.Value = True
         
      End If
      Me.txtMEAL.Value = Me.bukadatabase.List(Me.bukadatabase.ListIndex, 4)
      Me.txtPRICE.Value = Me.bukadatabase.List(Me.bukadatabase.ListIndex, 5)
      Me.txtAmt.Value = Me.bukadatabase.List(Me.bukadatabase.ListIndex, 6)
            
      MsgBox "Please make required changes and save to update", vbOKOnly + vbInformation, "Edit"
      
      
End Sub

Private Sub EXIT1_Click()
Unload Me
End Sub

Private Sub txtPrint_Click()
    Dim PrintRange As Range
    Set PrintRange = Selection
    
    If PrintRange Is Nothing Then
        MsgBox "Please select a range to print.", vbExclamation, "Print Selection"
    Else
        PrintRange.PrintOut
    End If
End Sub



Private Sub txtRESET_Click()

   Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Do you want to reset the form?", vbYesNo + vbInformation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call Reset
     
    
End Sub

Private Sub txtSAVE_Click()
    Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Do you want to save the data?", vbYesNo + vbInformation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call Submit
    Call Reset
    
    
End Sub


Private Sub UserForm_Initialize()
    
    Call Reset
    
End Sub

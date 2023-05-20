Attribute VB_Name = "Module7"
Option Explicit
Sub Show_AdminCenter()
    AdminCenter.Show
End Sub

Function UserExists(user As String) As Boolean
    Dim userSheet As Worksheet
    Set userSheet = ThisWorkbook.Sheets("Users")
    
    Dim lastRow As Long
    lastRow = userSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If userSheet.Cells(i, "A").Value = user Then
            UserExists = True
            Exit Function
        End If
    Next i
    
    UserExists = False
End Function

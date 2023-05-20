Attribute VB_Name = "Module5"
Option Explicit

Sub Redo()
    Dim iRow As Long
    iRow = [Counta(InventoryTesting!A:A)] ' identifying the last row
     With INVFORM
     .cmbAdded.Value = ""
     
     
.cmbCategory1.Clear
.cmbCategory1.AddItem "spices"
.cmbCategory1.AddItem "Seasoning"
.cmbCategory1.AddItem "Fruits"
.cmbCategory1.AddItem "Grain"
.cmbCategory1.AddItem "Vegetables"
.cmbCategory1.AddItem "Tuber"
.cmbCategory1.AddItem "Oils"
.cmbCategory1.AddItem "Peas"

     
     
.cmbIngredient.Clear
.cmbIngredient.AddItem "African eggplant"
 .cmbIngredient.AddItem "African Nutmeg"
 .cmbIngredient.AddItem "Alligator pepper"
.cmbIngredient.AddItem " Avocado"
.cmbIngredient.AddItem " Baobab fruit"
.cmbIngredient.AddItem " Beans"
.cmbIngredient.AddItem " Berbere spice mix"
.cmbIngredient.AddItem " Bitter leaf"
.cmbIngredient.AddItem " Black pepper"
.cmbIngredient.AddItem " Cardamom"
 .cmbIngredient.AddItem "Cassava"
 .cmbIngredient.AddItem "Chilies"
 .cmbIngredient.AddItem "Cinnamon"
 .cmbIngredient.AddItem "Cloves"
 .cmbIngredient.AddItem "Coconut oil"
 .cmbIngredient.AddItem "Coriander"
 .cmbIngredient.AddItem "Couscous"
 .cmbIngredient.AddItem "Cumin"
.cmbIngredient.AddItem " Egusi"
 .cmbIngredient.AddItem "Fonio"
 .cmbIngredient.AddItem "Garlic"
 .cmbIngredient.AddItem "Ginger"
 .cmbIngredient.AddItem "Grains of paradise"
 .cmbIngredient.AddItem "Mango"
 .cmbIngredient.AddItem "Millet"
.cmbIngredient.AddItem " Njangsa"
.cmbIngredient.AddItem " Nutmeg"
 .cmbIngredient.AddItem "Ogbono"
 .cmbIngredient.AddItem "Okra"
 .cmbIngredient.AddItem "Onions"
 .cmbIngredient.AddItem "Palm oil"
 .cmbIngredient.AddItem "Paprika"
 .cmbIngredient.AddItem "Peanut oil"
 .cmbIngredient.AddItem "Peanuts"
 .cmbIngredient.AddItem "Pineapple"
 .cmbIngredient.AddItem "Plantain"
.cmbIngredient.AddItem " Rice"
 .cmbIngredient.AddItem "Salt"
 .cmbIngredient.AddItem "Sesame oil"
 .cmbIngredient.AddItem "Shea butter"
.cmbIngredient.AddItem " Sorghum"
.cmbIngredient.AddItem " Suya spice mix"
.cmbIngredient.AddItem " Sweet potatoes"
 .cmbIngredient.AddItem "Tamarind"
 .cmbIngredient.AddItem "Tomatoes"
 .cmbIngredient.AddItem "Turmeric"
 .cmbIngredient.AddItem "Uziza leaf"
 .cmbIngredient.AddItem "Vegetable oil"
 .cmbIngredient.AddItem "Water leaf"
 .cmbIngredient.AddItem "Yams"

     
     
     .cmbAdded.Clear
 .cmbAdded.AddItem "1"
.cmbAdded.AddItem "2"
.cmbAdded.AddItem "3"
.cmbAdded.AddItem "4"
.cmbAdded.AddItem "5"
.cmbAdded.AddItem "6"
.cmbAdded.AddItem "7"
.cmbAdded.AddItem "8"
.cmbAdded.AddItem "9"
.cmbAdded.AddItem "10"
.cmbAdded.AddItem "11"
.cmbAdded.AddItem "12"
.cmbAdded.AddItem "13"
.cmbAdded.AddItem "14"
.cmbAdded.AddItem "15"
.cmbAdded.AddItem "16"
.cmbAdded.AddItem "17"
.cmbAdded.AddItem "18"
.cmbAdded.AddItem "19"
.cmbAdded.AddItem "20"
.cmbAdded.AddItem "21"
.cmbAdded.AddItem "22"
.cmbAdded.AddItem "23"
.cmbAdded.AddItem "24"
.cmbAdded.AddItem "25"
.cmbAdded.AddItem "26"
.cmbAdded.AddItem "27"
.cmbAdded.AddItem "28"
.cmbAdded.AddItem "29"
.cmbAdded.AddItem "30"
 
     
    
    .cmbUsed.Clear
    .cmbUsed.AddItem "1"
.cmbUsed.AddItem "2"
.cmbUsed.AddItem "3"
.cmbUsed.AddItem "4"
.cmbUsed.AddItem "5"
.cmbUsed.AddItem "6"
.cmbUsed.AddItem "7"
.cmbUsed.AddItem "8"
.cmbUsed.AddItem "9"
.cmbUsed.AddItem "10"
.cmbUsed.AddItem "11"
.cmbUsed.AddItem "12"
.cmbUsed.AddItem "13"
.cmbUsed.AddItem "14"
.cmbUsed.AddItem "15"
.cmbUsed.AddItem "16"
.cmbUsed.AddItem "17"
.cmbUsed.AddItem "18"
.cmbUsed.AddItem "19"
.cmbUsed.AddItem "20"
.cmbUsed.AddItem "21"
.cmbUsed.AddItem "22"
.cmbUsed.AddItem "23"
.cmbUsed.AddItem "24"
.cmbUsed.AddItem "25"
.cmbUsed.AddItem "26"
.cmbUsed.AddItem "27"
.cmbUsed.AddItem "28"
.cmbUsed.AddItem "29"
.cmbUsed.AddItem "30"




.cmbCosts.Clear
.cmbCosts.AddItem "3.2"
.cmbCosts.AddItem "2.5"
.cmbCosts.AddItem "2.2"
.cmbCosts.AddItem "1.5"
.cmbCosts.AddItem "3.5"
.cmbCosts.AddItem "3.1"
.cmbCosts.AddItem "2.3"
.cmbCosts.AddItem "1.8"
.cmbCosts.AddItem "1.2"
.cmbCosts.AddItem "2"
.cmbCosts.AddItem "1.4"
.cmbCosts.AddItem "2.1"
.cmbCosts.AddItem "3.3"
.cmbCosts.AddItem "1.6"
.cmbCosts.AddItem "1"
.cmbCosts.AddItem "1.9"
.cmbCosts.AddItem "1.1"
.cmbCosts.AddItem "2.7"
.cmbCosts.AddItem "2.8"
.cmbCosts.AddItem "1.3"
.cmbCosts.AddItem "1.7"
.cmbCosts.AddItem "2.9"
.cmbCosts.AddItem "2.6"
.cmbCosts.AddItem "2.4"
.cmbCosts.AddItem "3.4"


     
     .txtrowno.Value = ""
     .cmbAdded.Value = ""
     
     
     .Invdatabase.ColumnCount = 7
     .Invdatabase.ColumnHeads = True
     .Invdatabase.ColumnWidths = "35,60,75,60,60,60,70"
     
     If iRow > 1 Then
         .Invdatabase.RowSource = "InventoryTesting!A2:J" & iRow
     Else
         .Invdatabase.RowSource = "InventoryTesting!A2:J2"
     End If
 
End With

 
 
End Sub
Sub Send()

Dim sh As Worksheet
Dim iRow As Long

Set sh = ThisWorkbook.Sheets("InventoryTesting")

If INVFORM.txtrowno.Value = "" Then

   iRow = [Counta(InventoryTesting!A:A)] + 1
Else
   iRow = INVFORM.txtrowno.Value
   
End If

With sh

    .Cells(iRow, 1) = iRow - 1
    .Cells(iRow, 7) = INVFORM.cmbAdded.Value
    .Cells(iRow, 3) = INVFORM.cmbCosts.Value
    .Cells(iRow, 6) = INVFORM.cmbUsed.Value
    .Cells(iRow, 5) = [Text(Now(),"DD-MM-YYYY-HH:MM:SS")]
    .Cells(iRow, 4) = INVFORM.cmbCategory1.Value
    .Cells(iRow, 2) = INVFORM.cmbIngredient.Value
   
    
    
End With

End Sub


Sub display_Form()
  
    INVFORM.Show
     
End Sub

Function selected_form() As Long

    Dim i As Long
    
    selected_form = 0
    
    For i = 0 To INVFORM.Invdatabase.ListCount - 1
    
         If INVFORM.Invdatabase.Selected(i) = True Then
         
            selected_form = i + 1
            
            Exit For
            
         End If
         
    Next i
    
    
End Function



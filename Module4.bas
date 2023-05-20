Attribute VB_Name = "Module4"

Public logininstance As Integer

Sub Reset()
    Dim iRow As Long
    iRow = [Counta(Database!A:A)] ' identifying the last row
     With FrmForm
     .txtID.Value = ""
     .txtNAME.Value = ""
     .optMALE.Value = False
     .optFEMALE.Value = False
     
     .cmbCATEGORY.Clear
     .cmbCATEGORY.AddItem "STEW/ TROSKINYS"
     .cmbCATEGORY.AddItem "HOT DISHES/KARSTI PATIKELAI"
     .cmbCATEGORY.AddItem "GARNISHES I GARNYRAI"
     .cmbCATEGORY.AddItem "SIDES PUSÉS"
     .cmbCATEGORY.AddItem "SNACKS I UZKANDZIAI"
     .cmbCATEGORY.AddItem "DESERT / DESERTAI:"
     .cmbCATEGORY.AddItem "NON - ALCOHOL DRINKS / NEALKOHOLINIAI GÉRIMAI:"
     
     .txtMEAL.Clear
     .txtMEAL.AddItem "Pounded yam / jamso milty kose with Okro or(Okra)"
     .txtMEAL.AddItem "Pounded yam / jamso milty kose With Efo Riro"
     .txtMEAL.AddItem "Semolina / many kose with Ogbono"
     .txtMEAL.AddItem "Semolina / many kose with Okro or(Okra)"
     .txtMEAL.AddItem "Amala(yam flour)/jamo miltai With Egusi"
     .txtMEAL.AddItem "Fufu with Egusi"
     .txtMEAL.AddItem "Garri(eba) / fermentuota malta kasava With Efo Riro"
     .txtMEAL.AddItem "Semolina / many kose With Efo Riro"
     .txtMEAL.AddItem "White rice With Efo Riro"
     .txtMEAL.AddItem " Garri(eba) / fermentuota malta kasava With Egusi"
     .txtMEAL.AddItem "Pounded yam / jamso milty kose Egusi"
     .txtMEAL.AddItem "Semolina / many kose with Egusi"
     .txtMEAL.AddItem "Fried/fish(mackerel Fried stew)/ Keptazuvis (skumbre"
     .txtMEAL.AddItem "Goat meat Fried stew / OZkiena"
     .txtMEAL.AddItem "Beef Fried stew/ Jautiena"
     .txtMEAL.AddItem "Stew goat meat with white rice + dodo /Troskinta ozkiena su baltais ryziais,(ozkiena, planteinas,paprika, pomidorai, aliejus, svogúnai, kmynai, karis)"
     .txtMEAL.AddItem "Fufu with Ogbono"
     .txtMEAL.AddItem "Amala(yam flour)/Jamo miltai with Ogbono"
     .txtMEAL.AddItem "Pounded yam / jamso milty kose with Ogbono"
     .txtMEAL.AddItem "Amala(yam flour)/ jamo miltai With Efo Riro"
     .txtMEAL.AddItem "Amala(yam flour)/Jamo miltai with Okro or(Okra)"
     .txtMEAL.AddItem " Fufu with Okro or(Okra)"
     .txtMEAL.AddItem "Garri(eba) / fermentuota malta kasava with Ogbono"
     .txtMEAL.AddItem "Garri(eba) / fermentuota malta kasava with Okro or(Okra)"
    .txtMEAL.AddItem "Potato stew / (bulvés, morkos, svogünai, pomidorai, cesnakai,, paprika, aliejus, vistienos sultinys, pipirai, karis,vistiena)"
    .txtMEAL.AddItem "Potato porridge / Spinatai ,Darzoviy aliejus,, Raudonosios paprikos, pomidorai, svogúnas, imbiero cesnako, karis, skumbrê)"
    .txtMEAL.AddItem "BBQ chicken wings,(spicy / nospicy) (astrus/ neastrüs),(barbekiu vistienos sparneliai))"
    .txtMEAL.AddItem "Chicken Nuggets with french fries / Vistienos gabaléliai su bulvytèmis fri."
    .txtMEAL.AddItem "BBQ turkey wings (spicy / nospicy) (astrus/ neastrüs) / (barbekiu kalakutienos sparnelis)"
    .txtMEAL.AddItem "Fried rice with beef"
    .txtMEAL.AddItem "White Rice / Balti ryziai"
    .txtMEAL.AddItem "Spaghetti with veggies / Spagediai, (spicy / nospicy) (astrüs/ neastrüs), (spageciai, morkos, zirneliai, paprikos, svogünai, ¿esnakai, pipirai, palmiy aliejus)"
    .txtMEAL.AddItem "Spaghetti Jollof / Spageciai / (spicy / nospicy) (astrus/ neastrüs) / (spageciai, morkos, Zirneliai, paprikos, svogúnai, cesnakai, pipirai, palmiy aliejus) /"
   .txtMEAL.AddItem " Fried dodo with eggs/ Keptas dodo sukiausiniais"
    .txtMEAL.AddItem "Dodo / Kepti planteino griezinéliai (planteinas, aliejus, druska) (8 vnt)"
    .txtMEAL.AddItem "Beans Dodo with fried stew (baltosios pupeles, palmiy allejus, svogunal, vistienos sultinys, paprika, planteinas)"
    .txtMEAL.AddItem "Moin moin / Garintas pupeliu pyragas (pupeliu miltai, vistienos sultinys, paprika, svogünai, skumbre, kiausiniai, aliejus)"
    .txtMEAL.AddItem "Plantain chips / Planteiny traskuciai (jautiena, bulvés, morkos, svogüny laiskai, svogünai, kviediu 4,00 € miltai, aliejus, kmynai, jautienos sultinys, cesnakai)"
   .txtMEAL.AddItem " Meat pie / Mésos pyragas"
    .txtMEAL.AddItem "Puff Puff / saldzios spurgytes (8 vnt.)"
    .txtMEAL.AddItem "Onion rings / Svogüny Ziedai (8 vnt.)"
    .txtMEAL.AddItem "Ice cream with fruit / ledai su vaisiais"
    .txtMEAL.AddItem "Waffles"
    .txtMEAL.AddItem "Sprite (0,250 l)"
    .txtMEAL.AddItem "Fanta (0,250 l)"
    .txtMEAL.AddItem "Coca - cola / Coca - Cola zero (0,250 l)"
   .txtMEAL.AddItem " Cheesecake / sûrio pyragas"
   .txtMEAL.AddItem " Schweppes (0,250 l)"
    .txtMEAL.AddItem "Mineral water (carbonated, still) / Mineralinis vanduo (gazuotas, negazuotas) (0,330 l/0,750 l)"
    .txtMEAL.AddItem "Maltina Guinness"
    .txtMEAL.AddItem "Jollof rice with beef +dodo"
    .txtMEAL.AddItem "Jollof rice with beef"
    .txtMEAL.AddItem "Jollof Rice / Ryziai (spicy / nospicy) (astrüs/ neastrüs), (ryziai, vistienos sultinys, morkos, zirneliai, paprika, svogúnai,pomidorai)"
   .txtMEAL.AddItem " Indomie with Fried Eggs ('Indomie' makaronai su kiausiniu ir darzovémis (morkos, zirneliai, paprika, karis, aliejus)"
    .txtMEAL.AddItem "Fried rice with fish+dodo"
    .txtMEAL.AddItem "Fried Rice / Ryziai, (spicy / nospicy) (astrüs/ neastrüs) / (ryziai, ciberzole, zirneliai, morkos, paprika, jautiena, vistienos sultinys, svogunai)"
    .txtMEAL.AddItem "Fried Rice / Ryziai with Goat meat/ Ozkiena / (spicy / nospicy) (astrüs/ neastrüs) / (ryziai, ciberzole, zirneliai, morkos, paprika, jautiena, vistienos sultinys, svogünai)"
    .txtMEAL.AddItem "Fried Rice / Ryziai with Fried fish(mackerel) / Keptazuvis (spicy / nospicy) (astrüs/ neastrüs)/ (skumbrê)"
    .txtMEAL.AddItem "Fried fish with stew (mackerel) / Kepta zuvis (skumbre). (kepta zuvis marinuota zuvies prieskoniuose)"
    .txtMEAL.AddItem "Moin moin / Garintas pupeliy pyragas / (pupeliy miltai, vistienos sultinys, paprika, svogúnai, skumbre, kiausiniai, aliejus)"
    .txtMEAL.AddItem "Fried beef with stew / Kepta jautiena su troskiniu."
    
    .txtPRICE.Clear
    .txtPRICE.AddItem "9.5"
 .txtPRICE.AddItem "11.5"
  .txtPRICE.AddItem "10.5"
 .txtPRICE.AddItem "9.5"
 .txtPRICE.AddItem "9"
 .txtPRICE.AddItem "5.5"
 .txtPRICE.AddItem "6.5"
.txtPRICE.AddItem " 8.5"
.txtPRICE.AddItem "7.5"
.txtPRICE.AddItem "4.5"
.txtPRICE.AddItem "12.5"
.txtPRICE.AddItem "5"
 .txtPRICE.AddItem "4"
 .txtPRICE.AddItem "6"
 .txtPRICE.AddItem "2"
 .txtPRICE.AddItem "2.5"
 .txtPRICE.AddItem "14"
 .txtPRICE.AddItem "11"
.txtPRICE.AddItem "15"


.txtDIS.Clear
.txtDIS.AddItem "0.25"
.txtDIS.AddItem "0.22"
.txtDIS.AddItem "0.25"
.txtDIS.AddItem "0.35"
.txtDIS.AddItem "0.15"
.txtDIS.AddItem "0.2"
.txtDIS.AddItem "0.1"
.txtDIS.AddItem "0.14"
.txtDIS.AddItem "0.12"
.txtDIS.AddItem "0.16"
.txtDIS.AddItem "0.33"
.txtDIS.AddItem "0.21"
.txtDIS.AddItem "0.18"
.txtDIS.AddItem "0.19"
.txtDIS.AddItem "0.11"
.txtDIS.AddItem "0.27"
.txtDIS.AddItem "0.28"
.txtDIS.AddItem "0.13"
.txtDIS.AddItem "0.17"
.txtDIS.AddItem "0.29"
.txtDIS.AddItem "0.26"
.txtDIS.AddItem "0.24"
.txtDIS.AddItem "0.34"
.txtDIS.AddItem "0.3"

.txtAmt.Clear
.txtAmt.AddItem "1"
.txtAmt.AddItem "2"
.txtAmt.AddItem "3"
.txtAmt.AddItem "4"
.txtAmt.AddItem "5"
.txtAmt.AddItem "6"
.txtAmt.AddItem "7"
.txtAmt.AddItem "8"
.txtAmt.AddItem "9"
.txtAmt.AddItem "10"
.txtAmt.AddItem "11"
.txtAmt.AddItem "12"
.txtAmt.AddItem "13"
.txtAmt.AddItem "14"
.txtAmt.AddItem "15"
.txtAmt.AddItem "16"
.txtAmt.AddItem "17"
.txtAmt.AddItem "18"
.txtAmt.AddItem "19"
.txtAmt.AddItem "20"

.txtID.Clear
.txtID.AddItem "OO1"
.txtID.AddItem "OO2"
.txtID.AddItem "OO4"
.txtID.AddItem "OO5"
.txtID.AddItem "OO6"
.txtID.AddItem "OO7"
.txtID.AddItem "OO8"

     
     .txtROWNUMBER.Value = ""
     .txtMEAL.Value = ""
     .txtAmt.Value = ""
     
     
     .bukadatabase.ColumnCount = 10
     .bukadatabase.ColumnHeads = True
     
     .bukadatabase.ColumnWidths = "35,60,75,60,60,60,70,70,70,70"
     
     If iRow > 1 Then
         .bukadatabase.RowSource = "Database!A2:J" & iRow
     Else
         .bukadatabase.RowSource = "Database!A2:J2"
     End If
 
End With

 
 
End Sub
Sub Submit()

Dim sh As Worksheet
Dim iRow As Long

Set sh = ThisWorkbook.Sheets("Database")

If FrmForm.txtROWNUMBER.Value = "" Then

   iRow = [Counta(Database!A:A)] + 1
Else
   iRow = FrmForm.txtROWNUMBER.Value
   
End If

With sh

    .Cells(iRow, 1) = iRow - 1
    .Cells(iRow, 2) = FrmForm.txtID.Value
    .Cells(iRow, 3) = FrmForm.txtNAME.Value
    .Cells(iRow, 4) = IIf(FrmForm.optFEMALE.Value = True, "Female", "Male")
    .Cells(iRow, 5) = FrmForm.txtMEAL.Value
    .Cells(iRow, 6) = [Text(Now(),"DD-MM-YYYY-HH:MM:SS")]
    .Cells(iRow, 7) = FrmForm.cmbCATEGORY.Value
    .Cells(iRow, 8) = FrmForm.txtAmt.Value
    .Cells(iRow, 9) = FrmForm.txtDIS.Value
    .Cells(iRow, 10) = FrmForm.txtPRICE.Value
    
    
    
End With

End Sub


Sub Show_Form()
  
    FrmForm.Show
     
End Sub

Function selected_list() As Long

    Dim i As Long
    
    selected_list = 0
    
    For i = 0 To FrmForm.bukadatabase.ListCount - 1
    
         If FrmForm.bukadatabase.Selected(i) = True Then
         
            selected_list = i + 1
            
            Exit For
            
         End If
         
    Next i
    
    
End Function


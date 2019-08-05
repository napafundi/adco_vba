
Private Sub Cancel2_Click()

Unload Me

End Sub

Private Sub Enter2_Click()

'check the selected product and index the cell value within the sheet to edit

'ironweed bottles
If product_list.Value = "Louisville Bottles" Then
material_index = WorksheetFunction.Match("Louisville (Ironweed)", Sheets("Bottles").Range("A:A"), 0)
sheet_index = "Bottles"


'alb 1L bottles
ElseIf product_list.Value = "Straight Up 1L (ALB Vodka) Bottles" Then
material_index = WorksheetFunction.Match("Straight Up 1L (ALB Vodka)", Sheets("Bottles").Range("A:A"), 0)
sheet_index = "Bottles"

'amber rum 750ml bottles
ElseIf product_list.Value = "Straight Up 750 (Amber Rum) Bottles" Then
material_index = WorksheetFunction.Match("Straight Up 750 (Amber Rum)", Sheets("Bottles").Range("A:A"), 0)
sheet_index = "Bottles"


'death wish bottles
ElseIf product_list.Value = "Death Wish Bottles" Then
material_index = WorksheetFunction.Match("Death Wish Bottle", Sheets("Bottles").Range("A:A"), 0)
sheet_index = "Bottles"


'white rum bottles
ElseIf product_list.Value = "White Rum Bottles" Then
material_index = WorksheetFunction.Match("White Rum", Sheets("Bottles").Range("A:A"), 0)
sheet_index = "Bottles"


'200mL bottles
ElseIf product_list.Value = "200ml Bottles" Then
material_index = WorksheetFunction.Match("200ml", Sheets("Bottles").Range("A:A"), 0)
sheet_index = "Bottles"


'50ml bottles
ElseIf product_list.Value = "50ml Bottles" Then
material_index = WorksheetFunction.Match("50ml", Sheets("Bottles").Range("A:A"), 0)
sheet_index = "Bottles"


'alb mag bottles
ElseIf product_list.Value = "ALB Mag Bottles" Then
material_index = WorksheetFunction.Match("ALB Mag", Sheets("Bottles").Range("A:A"), 0)
sheet_index = "Bottles"


'ironweed boxes
ElseIf product_list.Value = "Ironweed Boxes" Then
material_index = WorksheetFunction.Match("Ironweed", Sheets("Boxes").Range("A:A"), 0)
sheet_index = "Boxes"


'straight up 750 boxes
ElseIf product_list.Value = "Straight Up 750 Boxes" Then
material_index = WorksheetFunction.Match("Straight Up 750", Sheets("Boxes").Range("A:A"), 0)
sheet_index = "Boxes"


'death wish boxes
ElseIf product_list.Value = "Death Wish Boxes" Then
material_index = WorksheetFunction.Match("Death Wish", Sheets("Boxes").Range("A:A"), 0)
sheet_index = "Boxes"


'50ml boxes
ElseIf product_list.Value = "50ml Boxes" Then
material_index = WorksheetFunction.Match("50ml", Sheets("Boxes").Range("A:A"), 0)
sheet_index = "Boxes"

'ALB Mag boxes
ElseIf product_list.Value = "ALB Mag Boxes" Then
material_index = WorksheetFunction.Match("ALB Mag", Sheets("Boxes").Range("A:A"), 0)
sheet_index = "Boxes"


'ironweed caps
ElseIf product_list.Value = "Ironweed Caps" Then
material_index = WorksheetFunction.Match("Ironweed", Sheets("Caps").Range("A:A"), 0)
sheet_index = "Caps"


'quackenbush amber caps
ElseIf product_list.Value = "Quackenbush Amber Caps" Then
material_index = WorksheetFunction.Match("Quackenbush Amber", Sheets("Caps").Range("A:A"), 0)
sheet_index = "Caps"


'33x10/22.7mm (DW/White Rum)
ElseIf product_list.Value = "33x10/22.7mm (DW/White Rum) Caps" Then
material_index = WorksheetFunction.Match("33x10/22.7mm (DW/White Rum)", Sheets("Caps").Range("A:A"), 0)
sheet_index = "Caps"


ElseIf product_list.Value = "29x10/19.5mm (ALB) Caps" Then
material_index = WorksheetFunction.Match("29x10/19.5mm (ALB)", Sheets("Caps").Range("A:A"), 0)
sheet_index = "Caps"


'200ml caps
ElseIf product_list.Value = "200ml Caps" Then
material_index = WorksheetFunction.Match("200ml", Sheets("Caps").Range("A:A"), 0)
sheet_index = "Caps"


'50ml caps
ElseIf product_list.Value = "50ml Caps" Then
material_index = WorksheetFunction.Match("50ml", Sheets("Caps").Range("A:A"), 0)
sheet_index = "Caps"


'ALB Mag caps
ElseIf product_list.Value = "ALB Mag Caps" Then
material_index = WorksheetFunction.Match("ALB Mag", Sheets("Caps").Range("A:A"), 0)
sheet_index = "Caps"


'Ironweed capsules
ElseIf product_list.Value = "Ironweed/Quack White Capsules" Then
material_index = WorksheetFunction.Match("Ironweed/Quack White", Sheets("Capsules").Range("A:A"), 0)
sheet_index = "Capsules"


'death wish capsules
ElseIf product_list.Value = "Death Wish Capsules" Then
material_index = WorksheetFunction.Match("Death Wish", Sheets("Capsules").Range("A:A"), 0)
sheet_index = "Capsules"


'straight up 750 capsules
ElseIf product_list.Value = "Straight Up 750 Capsules" Then
material_index = WorksheetFunction.Match("Straight Up 750", Sheets("Capsules").Range("A:A"), 0)
sheet_index = "Capsules"


'ALB/Pride/Fort O capsules
ElseIf product_list.Value = "ALB/Pride/Fort O Capsules" Then
material_index = WorksheetFunction.Match("ALB/Pride/Fort O", Sheets("Capsules").Range("A:A"), 0)
sheet_index = "Capsules"


'200ml seal capsules
ElseIf product_list.Value = "200ml Seal Capsules" Then
material_index = WorksheetFunction.Match("200ml Seal", Sheets("Capsules").Range("A:A"), 0)
sheet_index = "Capsules"

'ALB Mag capsules
ElseIf product_list.Value = "ALB Mag Capsules" Then
material_index = WorksheetFunction.Match("ALB Mag", Sheets("Capsules").Range("A:A"), 0)
sheet_index = "Capsules"


'ironweed rye 750 labels
ElseIf product_list.Value = "Ironweed Rye 750 Labels" Then
material_index = WorksheetFunction.Match("Ironweed Rye 750", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'ironweed bourbon 750 labels
ElseIf product_list.Value = "Ironweed Bourbon 750 Labels" Then
material_index = WorksheetFunction.Match("Ironweed Bourbon 750", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"



'ironweed malt 750 labels
ElseIf product_list.Value = "Ironweed Malt 750 Labels" Then
material_index = WorksheetFunction.Match("Ironweed Malt 750", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'ironweed rye 200 labels
ElseIf product_list.Value = "Ironweed Rye 200 Labels" Then
material_index = WorksheetFunction.Match("Ironweed Rye 200", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'ironweed bourbon 200 labels
ElseIf product_list.Value = "Ironweed Bourbon 200 Labels" Then
material_index = WorksheetFunction.Match("Ironweed Bourbon 200", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'ironweed malt 200 labels
ElseIf product_list.Value = "Ironweed Malt 200 Labels" Then
material_index = WorksheetFunction.Match("Ironweed Malt 200", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'alb 1L labels
ElseIf product_list.Value = "ALB Vodka 1L Labels" Then
material_index = WorksheetFunction.Match("ALB Vodka 1L", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'alb 200 labels
ElseIf product_list.Value = "ALB Vodka 200 Labels" Then
material_index = WorksheetFunction.Match("ALB Vodka 200", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'death wish 50mL labels
ElseIf product_list.Value = "Death Wish 50ml Labels" Then
material_index = WorksheetFunction.Match("Death Wish 50ml", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'amber rum labels
ElseIf product_list.Value = "Amber Rum Labels" Then
material_index = WorksheetFunction.Match("Amber Rum", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'white rum labels
ElseIf product_list.Value = "White Rum Labels" Then
material_index = WorksheetFunction.Match("White Rum", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'pride labels
ElseIf product_list.Value = "Pride Labels" Then
material_index = WorksheetFunction.Match("Pride", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'fort orange labels
ElseIf product_list.Value = "Fort Orange Labels" Then
material_index = WorksheetFunction.Match("Fort Orange", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


'alb mag labels
ElseIf product_list.Value = "ALB Mag Labels" Then
material_index = WorksheetFunction.Match("ALB Mag", Sheets("Labels").Range("A:A"), 0)
sheet_index = "Labels"


End If

'Check to see if product type has been selected
If IsEmpty(material_index) Then
    MsgBox "You need to select a product type!", vbOKOnly, "Attention!"

End If

'check to see if an amount has been entered
If Len(Me.amount_box2 & vbNullString) = 0 Then
    MsgBox "You need to enter an amount!", vbOKOnly, "Attention!"

End If

'Check to see if all selections have been made, then edits the corresponding inventory values and closes the userform
If IsEmpty(material_index) = False And Len(Me.amount_box2 & vbNullString) <> 0 Then
    
    Sheets(sheet_index).Range("c" & material_index) = Sheets(sheet_index).Range("c" & material_index) + amount_box2.Value

    
 'Add information to "raw materials received" worksheet
    Set ws = Sheets("Raw Materials Received")
    Set tbl = ws.ListObjects("raw_materials_table")
    Set newrow = tbl.ListRows.Add
    
    With newrow
        .Range("1") = Date
        .Range("2") = product_list.Value
        .Range("3") = amount_box2.Value
        .Range("4") = notes_box2.Value
    End With

Unload Me

End If




End Sub

Private Sub notes_box2_Change()

End Sub


Private Sub UserForm_Initialize()

'Initalize some variables

Dim material_index As Integer
Dim sheet_index As Integer
Dim ws2 As Worksheet
Dim tbl2 As ListObject
Dim newrow2 As ListRow

Dim Bottles_Array_2

'Fill Product_List Box with the company product names

With product_list
    Bottles_Array_2 = Array("Louisville Bottles", "Straight Up 1L (ALB Vodka) Bottles", "Straight Up 750 (Amber Rum) Bottles", "Death Wish Bottles", "White Rum Bottles", "200ml Bottles", "50ml Bottles", "ALB Mag Bottles", "Ironweed Boxes", "Straight Up 750 Boxes", "Death Wish Boxes", "50ml Boxes", "ALB Mag Boxes", "Ironweed Caps", "Quackenbush Amber Caps", "33x10/22.7mm (DW/White Rum) Caps", "29x10/19.5mm (ALB) Caps", "200ml Caps", "50ml Caps", "ALB Mag Caps", "Ironweed/Quack White Capsules", "Death Wish Capsules", "Straight Up 750 Capsules", "ALB/Pride/Fort O Capsules", "200ml Seal Capsules", "ALB Mag Capsules", "Ironweed Rye 750 Labels", "Ironweed Bourbon 750 Labels", "Ironweed Malt 750 Labels", "Ironweed Rye 200 Labels", "Ironweed Bourbon 200 Labels", "Ironweed Malt 200 Labels", "ALB Vodka 1L Labels", "ALB Vodka 200 Labels", "Death Wish 50ml Labels", "Amber Rum Labels", "White Rum Labels", "Pride Labels", "Fort Orange Labels", "ALB Mag Labels")
    product_list.List = Bottles_Array_2
End With

End Sub
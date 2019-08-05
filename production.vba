Private Sub Cancel_Click()

Unload Me

End Sub


Private Sub Enter_Click()

'Assign values depending upon the bottle type selected
'Match function searches for the given string within the first column of the given sheet


'ALB 200mL
If bottles_list.Value = "ALB 200mL" Then
bottles_index = WorksheetFunction.Match("200ml", Sheets("Bottles").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("200ml", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("200ml Seal", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("ALB Vodka 200", Sheets("Labels").Range("A:A"), 0)

'ALB 1L
ElseIf bottles_list.Value = "ALB 1L" Then
bottles_index = WorksheetFunction.Match("Straight Up 1L (ALB Vodka)", Sheets("Bottles").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("29x10/19.5mm (ALB)", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("ALB/Pride/Fort O", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("ALB Vodka 1L", Sheets("Labels").Range("A:A"), 0)

'ALB 1.75L
ElseIf bottles_list.Value = "ALB 1.75L" Then
bottles_index = WorksheetFunction.Match("ALB Mag", Sheets("Bottles").Range("A:A"), 0)
boxes_index = WorksheetFunction.Match("ALB Mag", Sheets("Boxes").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("33x10/22.7mm (DW/White Rum)", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("ALB Mag", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("ALB Mag", Sheets("Labels").Range("A:A"), 0)

'ALB Pride 1L
ElseIf bottles_list.Value = "ALB Pride 1L" Then
bottles_index = WorksheetFunction.Match("Straight Up 1L (ALB Vodka)", Sheets("Bottles").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("29x10/19.5mm (ALB)", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("ALB/Pride/Fort O", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("Pride", Sheets("Labels").Range("A:A"), 0)

'ALB Fort Orange 1L
ElseIf bottles_list.Value = "ALB Fort Orange 1L" Then
bottles_index = WorksheetFunction.Match("Straight Up 1L (ALB Vodka)", Sheets("Bottles").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("29x10/19.5mm (ALB)", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("ALB/Pride/Fort O", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("Fort Orange", Sheets("Labels").Range("A:A"), 0)

'Death Wish 50mL
ElseIf bottles_list.Value = "Death Wish 50mL" Then
bottles_index = WorksheetFunction.Match("50ml", Sheets("Bottles").Range("A:A"), 0)
boxes_index = WorksheetFunction.Match("50ml", Sheets("Boxes").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("50ml", Sheets("Caps").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("Death Wish 50ml", Sheets("Labels").Range("A:A"), 0)

'Death Wish 1L
ElseIf bottles_list.Value = "Death Wish 1L" Then
bottles_index = WorksheetFunction.Match("Death Wish Bottle", Sheets("Bottles").Range("A:A"), 0)
boxes_index = WorksheetFunction.Match("Death Wish", Sheets("Boxes").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("33x10/22.7mm (DW/White Rum)", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("Death Wish", Sheets("Capsules").Range("A:A"), 0)


'Death Wish Cauldron
ElseIf bottles_list.Value = "Death Wish Cauldron" Then
bottles_index = WorksheetFunction.Match("Death Wish Bottle", Sheets("Bottles").Range("A:A"), 0)
boxes_index = WorksheetFunction.Match("Death Wish", Sheets("Boxes").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("33x10/22.7mm (DW/White Rum)", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("Death Wish", Sheets("Capsules").Range("A:A"), 0)

'Ironweed Bourbon 200mL
ElseIf bottles_list.Value = "Ironweed Bourbon 200mL" Then
bottles_index = WorksheetFunction.Match("200ml", Sheets("Bottles").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("200ml", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("200ml Seal", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("Ironweed Bourbon 200", Sheets("Labels").Range("A:A"), 0)

'Ironweed Bourbon 750mL
ElseIf bottles_list.Value = "Ironweed Bourbon 750mL" Then
bottles_index = WorksheetFunction.Match("Louisville (Ironweed)", Sheets("Bottles").Range("A:A"), 0)
boxes_index = WorksheetFunction.Match("Ironweed", Sheets("Boxes").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("Ironweed", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("Ironweed/Quack White", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("Ironweed Bourbon 750", Sheets("Labels").Range("A:A"), 0)

'Ironweed Rye 200mL
ElseIf bottles_list.Value = "Ironweed Rye 200mL" Then
bottles_index = WorksheetFunction.Match("200ml", Sheets("Bottles").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("200ml", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("200ml Seal", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("Ironweed Rye 200", Sheets("Labels").Range("A:A"), 0)

'Ironweed Rye 750mL
ElseIf bottles_list.Value = "Ironweed Rye 750mL" Then
bottles_index = WorksheetFunction.Match("Louisville (Ironweed)", Sheets("Bottles").Range("A:A"), 0)
boxes_index = WorksheetFunction.Match("Ironweed", Sheets("Boxes").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("Ironweed", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("Ironweed/Quack White", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("Ironweed Rye 750", Sheets("Labels").Range("A:A"), 0)

'Ironweed Malt 200mL
ElseIf bottles_list.Value = "Ironweed Malt 200mL" Then
bottles_index = WorksheetFunction.Match("200ml", Sheets("Bottles").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("200ml", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("200ml Seal", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("Ironweed Malt 200", Sheets("Labels").Range("A:A"), 0)

'Ironweed Malt 750mL
ElseIf bottles_list.Value = "Ironweed Malt 750mL" Then
bottles_index = WorksheetFunction.Match("Louisville (Ironweed)", Sheets("Bottles").Range("A:A"), 0)
boxes_index = WorksheetFunction.Match("Ironweed", Sheets("Boxes").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("Ironweed", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("Ironweed/Quack White", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("Ironweed Malt 750", Sheets("Labels").Range("A:A"), 0)


'Amber Rum 750mL
ElseIf bottles_list.Value = "Amber Rum 750mL" Then
bottles_index = WorksheetFunction.Match("Straight Up 750 (Amber Rum)", Sheets("Bottles").Range("A:A"), 0)
boxes_index = WorksheetFunction.Match("Straight Up 750", Sheets("Boxes").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("Quackenbush Amber", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("Straight Up 750", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("Amber Rum", Sheets("Labels").Range("A:A"), 0)

'White Rum 750mL
ElseIf bottles_list.Value = "White Rum 750mL" Then
bottles_index = WorksheetFunction.Match("White Rum", Sheets("Bottles").Range("A:A"), 0)
caps_index = WorksheetFunction.Match("33x10/22.7mm (DW/White Rum)", Sheets("Caps").Range("A:A"), 0)
capsules_index = WorksheetFunction.Match("Ironweed/Quack White", Sheets("Capsules").Range("A:A"), 0)
labels_index = WorksheetFunction.Match("White Rum", Sheets("Labels").Range("A:A"), 0)
End If

'Check to see if the amount box has been filled, if not return to the userform so that the user can enter an amount
If Len(Me.amount_box & vbNullString) = 0 Then
    MsgBox "You need to enter a value for bottle amount!", vbOKOnly, "Attention!"
    
End If

If IsEmpty(bottles_index) Then
    MsgBox "You need to select a bottle type!", vbOKOnly, "Attention!"

End If


'Check to see if all selections have been made, then edits the corresponding inventory values and closes the userform
If IsEmpty(bottles_index) = False And Len(Me.amount_box & vbNullString) <> 0 Then
    
    'bottles
    If Not IsEmpty(bottles_index) Then
    
    Sheets("Bottles").Range("c" & bottles_index) = Sheets("Bottles").Range("c" & bottles_index) - amount_box.Value
    
    End If
    
    'boxes
    If Not IsEmpty(boxes_index) Then
        
        'Check which bottle type is selected and subtract the correct number
        'of boxes per bottle
        If bottles_list.Value = "Death Wish 50mL" Then
        
        boxes_subtraction = Application.WorksheetFunction.RoundDown(amount_box.Value / 60, 0)
        
        ElseIf bottles_list.Value = "Death Wish 1L" Or bottles_list.Value = "Death Wish Cauldron" Then
        boxes_subtraction = Application.WorksheetFunction.RoundDown(amount_box.Value / 6, 0)
        
        ElseIf bottles_list.Value = "Ironweed Bourbon 750mL" Or bottles_list.Value = "Ironweed Rye 750mL" Or bottles_list.Value = "Ironweed Malt 750mL" Or bottles_list.Value = "Ironweed Malt 4yr" Then
        boxes_subtraction = Application.WorksheetFunction.RoundDown(amount_box.Value / 6, 0)
        
        ElseIf bottles_list.Value = "Amber Rum 750mL" Then
        boxes_subtraction = Application.WorksheetFunction.RoundDown(amount_box.Value / 6, 0)
        
        ElseIf bottles_list.Value = "ALB 1.75L" Then
        boxes_subtraction = Application.WorksheetFunction.RoundDown(amount_box.Value / 6, 0)
        
        End If
        
    Sheets("Boxes").Range("C" & boxes_index) = Sheets("Boxes").Range("C" & boxes_index) - boxes_subtraction
    
    End If
    
    'caps
    If Not IsEmpty(caps_index) Then
    
    Sheets("Caps").Range("c" & caps_index) = Sheets("Caps").Range("c" & caps_index) - amount_box.Value
    
    End If
    'capsules
    If Not IsEmpty(capsules_index) Then
    
    Sheets("Capsules").Range("c" & capsules_index) = Sheets("Capsules").Range("c" & capsules_index) - amount_box.Value
    
    End If
    
    'labels
    If Not IsEmpty(labels_index) Then
    
    Sheets("Labels").Range("c" & labels_index) = Sheets("Labels").Range("c" & labels_index) - amount_box.Value
    
    End If
    
     'Add information to "bottling log" worksheet
    Set ws = Sheets("Bottling Log")
    Set tbl = ws.ListObjects("bottling_log_table")
    Set newrow = tbl.ListRows.Add
    
    With newrow
        .Range("1") = Date
        .Range("2") = bottles_list.Value
        .Range("3") = amount_box.Value
        .Range("4") = notes_box.Value
    End With
    
    Unload Me
    
End If


End Sub

Private Sub UserForm_Initialize()


'Initalize some variables

Dim bottles_index As Integer
Dim cases_index As Integer
Dim caps_index As Integer
Dim capsules_index As Integer
Dim labels_index As Integer
Dim boxes_index As Integer
Dim boxes_subtraction As Integer
Dim ws As Worksheet
Dim tbl As ListObject
Dim newrow As ListRow



Dim Bottles_Array

'Fill Product_List Box with the company product names

With bottles_list
    Bottles_Array = Array("ALB 200mL", "ALB 1L", "ALB 1.75L", "ALB Pride 1L", "ALB Fort Orange 1L", "Death Wish 50mL", "Death Wish 1L", "Death Wish Cauldron", "Ironweed Bourbon 200mL", "Ironweed Bourbon 750mL", "Ironweed Rye 200mL", "Ironweed Rye 750mL", "Ironweed Malt 200mL", "Ironweed Malt 750mL", "Amber Rum 750mL", "White Rum 750mL")
    bottles_list.List = Bottles_Array
End With



End Sub

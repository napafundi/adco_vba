
Private Sub save_file_click()

Dim DateValue As Variant

' loops the inputbox until the user enters a nonempty value
Do
    DateValue = InputBox("Please enter a date value in the form: mm-dd-yyyy", "Date Confirmation")
    
    ' closes the inputbox if "cancel" is selected
    If StrPtr(DateValue) = 0 Then Exit Sub
    
    ' asks for another use input if nothing was entered
    If DateValue = "" Then MsgBox ("You must enter value in the form: mm-dd-yyyy")
    
    
Loop Until Not DateValue = ""

' explains what will occur when the user selects "yes"
msg_answer = MsgBox("Selecting 'Yes' will save this as a a new file within the 'Raw Materials Inventory' folder with the date attached", vbYesNo, "Raw Material Timestamp Creation")

If msg_answer = vbYes Then
    
    ' errors will go to the defined errHandler
    On Error GoTo errHandler:
    
    ' saves the workbook within the specified folder
    ActiveWorkbook.SaveAs ("C:\Users\napaf\Dropbox (ADCo)\ADCo Team Folder\Operations - ALL\Montgomery\Luc File\Inventory\Raw Materials Inventory\Raw Materials " & DateValue & ".xlsm")
    
    
    MsgBox "Your new raw material workbook has been saved within the 'Raw Materials Inventory' folder"
    
    ' closes the workbook so the user doesn't edit the incorrect file
    ActiveWorkbook.Close
Else
    
    Exit Sub
    
End If

'closes the inputbox if an error is reached
errHandler:

MsgBox ("There was an error in saving your file. Please try again and make sure you are entering the date format correctly.")
Exit Sub

End Sub

Private Sub raw_materials_click()

product_log.Show

End Sub

Sub production_entry_Click()

bottling_log.Show

End Sub
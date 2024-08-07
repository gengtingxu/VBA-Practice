Attribute VB_Name = "Module1"
Function Discount(Quantity, Price)
    If Quantity > 25 Then
        Discount = Quantity * Price * 0.2
    Else
        Discount = 0
    End If
    
End Function

Sub ClearContent()
    Answer = MsgBox("Confirm you want clear?", vbYesNo)
    Rows("6:" & Rows.Count).ClearContents
End Sub

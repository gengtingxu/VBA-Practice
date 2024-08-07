Attribute VB_Name = "Module1"
Sub Combine()
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name <> "Summary" Then
            j = Worksheets("Summary").Range("A10").End(xlUp).Row + 1
            Worksheets(i).UsedRange.Offset(1, 0).Copy Worksheets("Summary").Cells(j, 1)
        End If
    Next
    MsgBox "Comlete!"
End Sub

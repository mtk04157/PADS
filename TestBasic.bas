Sub Main
    Dim PCBDoc
    Set PCBDoc = ActiveDocument
    If PCBDoc Is Nothing Then
        MsgBox "No file"
    Else
        MsgBox "Success"
    End If
End Sub

Sub Main
    Dim PCBDoc
    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "PCB file not open"
    Else
        MsgBox "PCB file opened successfully"
    End If
End Sub

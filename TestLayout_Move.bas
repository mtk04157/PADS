Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim i

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    i = 0

    For Each comp In components
        If i = 0 Then
            comp.Move 5, 5
            MsgBox "Moved component", vbInformation
        End If
        i = i + 1
    Next comp

    PCBDoc.Refresh
    MsgBox "Complete", vbInformation
End Sub

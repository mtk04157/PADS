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
            MsgBox "Component: " & comp.Name, vbInformation
        End If
        i = i + 1
    Next comp

    MsgBox "Step 2 complete", vbInformation
End Sub

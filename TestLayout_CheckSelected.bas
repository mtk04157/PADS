Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim count

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    count = 0

    For Each comp In components
        If comp.selected Then
            count = count + 1
            MsgBox "Found selected: " & comp.Name, vbInformation
        End If
    Next comp

    If count = 0 Then
        MsgBox "No components selected", vbExclamation
    End If
End Sub

Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim compCount

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    compCount = 0

    For Each comp In components
        compCount = compCount + 1
    Next comp

    MsgBox "Found " & compCount & " components", vbInformation
End Sub

Sub Main
    On Error GoTo ErrorHandler

    Dim comp
    Dim i
    Dim compCount

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "Please open PCB file first", vbCritical
        Exit Sub
    End If

    compCount = 0
    i = 0

    For Each comp In PCBDoc.GetObjects(ppcbObjectTypeComponent,, True)
        comp.Move i * 2, 0
        i = i + 1
        compCount = compCount + 1
    Next comp

    On Error Resume Next
    PCBDoc.Refresh
    On Error GoTo 0

    MsgBox "Arranged " & compCount & " components", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

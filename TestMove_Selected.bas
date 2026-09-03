Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim components
    Dim comp
    Dim i
    Dim selectedCount

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    selectedCount = 0

    For Each comp In components
        If comp.selected Then
            selectedCount = selectedCount + 1
            On Error Resume Next
            comp.Move i * 1, 0
            On Error GoTo 0
            i = i + 1
        End If
    Next comp

    On Error Resume Next
    PCBDoc.Refresh
    On Error GoTo 0

    MsgBox "Moved " & selectedCount & " components", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

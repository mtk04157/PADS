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
    i = 0
    selectedCount = 0

    MsgBox "Starting component check", vbInformation

    For Each comp In components
        If comp.selected Then
            selectedCount = selectedCount + 1
            MsgBox "Component " & selectedCount & ": " & comp.Name, vbInformation
        End If
    Next comp

    MsgBox "Total selected: " & selectedCount, vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

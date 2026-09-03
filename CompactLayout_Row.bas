Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim components
    Dim comp
    Dim i
    Dim spacing

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "Please open PCB file first", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    spacing = 2
    i = 0

    For Each comp In components
        If comp.selected Then
            comp.Move i * spacing, 0
            i = i + 1
        End If
    Next comp

    On Error Resume Next
    PCBDoc.Refresh
    On Error GoTo 0

    MsgBox "Arranged " & i & " components in a row", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

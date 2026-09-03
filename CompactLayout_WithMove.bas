Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim components
    Dim comp
    Dim i

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "Please open PCB file first", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components

    i = 0
    For Each comp In components
        If comp.selected Then
            On Error Resume Next
            comp.Move i * 2, i * 2
            On Error GoTo 0
            i = i + 1
        End If
    Next comp

    On Error Resume Next
    PCBDoc.Refresh
    On Error GoTo 0

    If i = 0 Then
        MsgBox "No components selected"
    Else
        MsgBox "Moved " & i & " components"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

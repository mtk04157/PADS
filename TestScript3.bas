Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim x, y
    Dim i

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "PCB file not open"
        Exit Sub
    End If

    Set components = PCBDoc.Components

    i = 0
    For Each comp In components
        If comp.selected Then
            x = i * 1.0
            y = i * 1.0
            comp.LocationX = x
            comp.LocationY = y
            i = i + 1
        End If
    Next comp

    MsgBox "Moved " & i & " components"
    PCBDoc.Refresh
End Sub

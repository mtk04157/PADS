Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim width, height

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "PCB file not open"
        Exit Sub
    End If

    Set components = PCBDoc.Components

    For Each comp In components
        If comp.selected Then
            width = Abs(comp.BBox.Xmax - comp.BBox.Xmin)
            height = Abs(comp.BBox.Ymax - comp.BBox.Ymin)
            MsgBox "Component: " & comp.Name & " Width: " & width & " Height: " & height
            Exit For
        End If
    Next comp
End Sub

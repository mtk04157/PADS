Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim i
    Dim width, height

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    i = 0

    For Each comp In components
        If i = 0 Then
            width = Abs(comp.BBox.Xmax - comp.BBox.Xmin)
            height = Abs(comp.BBox.Ymax - comp.BBox.Ymin)
            MsgBox "Size: W=" & width & " H=" & height, vbInformation
        End If
        i = i + 1
    Next comp

    MsgBox "Step 4 complete", vbInformation
End Sub

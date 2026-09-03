Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim i

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    i = 0

    For Each comp In components
        If i = 0 Then
            comp.LocationX = 5
            comp.LocationY = 5
            MsgBox "Set position to 5,5", vbInformation
        End If
        i = i + 1
    Next comp

    PCBDoc.Refresh
    MsgBox "Step 5b complete", vbInformation
End Sub

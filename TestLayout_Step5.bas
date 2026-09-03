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
            comp.LocationX = 10
            comp.LocationY = 10
            MsgBox "Moved component to 10,10", vbInformation
        End If
        i = i + 1
    Next comp

    PCBDoc.Refresh
    MsgBox "Step 5 complete", vbInformation
End Sub

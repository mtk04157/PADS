Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim i
    Dim placement

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    i = 0

    For Each comp In components
        If i = 0 Then
            Set placement = comp.Placement
            placement.X = 5
            placement.Y = 5
            MsgBox "Set Placement object X,Y to 5,5", vbInformation
        End If
        i = i + 1
    Next comp

    PCBDoc.Refresh
    MsgBox "Complete", vbInformation
End Sub

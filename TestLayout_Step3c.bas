Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim i
    Dim x, y

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    i = 0

    For Each comp In components
        If i = 0 Then
            x = comp.XLocation
            y = comp.YLocation
            MsgBox "Position: X=" & x & " Y=" & y, vbInformation
        End If
        i = i + 1
    Next comp

    MsgBox "Step 3c complete", vbInformation
End Sub

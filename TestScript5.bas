Sub Main
    Dim PCBDoc
    Dim nets
    Dim net
    Dim netCount

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "PCB file not open"
        Exit Sub
    End If

    Set nets = PCBDoc.Nets
    netCount = 0

    For Each net In nets
        If net.PadCount >= 2 Then
            netCount = netCount + 1
        End If
    Next net

    MsgBox "Found " & netCount & " nets with 2+ pads"
End Sub

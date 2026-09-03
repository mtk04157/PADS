Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim nets
    Dim net
    Dim netCount

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set nets = PCBDoc.Nets
    netCount = 0

    For Each net In nets
        On Error Resume Next
        If net.PadCount >= 2 Then
            netCount = netCount + 1
        End If
        On Error GoTo 0
    Next net

    MsgBox "Found " & netCount & " nets with 2+ pads", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

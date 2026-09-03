Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim nets
    Dim net
    Dim i

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set nets = PCBDoc.Nets
    i = 0

    For Each net In nets
        On Error Resume Next
        If net.PadCount >= 2 And i = 0 Then
            MsgBox "Net: " & net.Name & " Pads: " & net.PadCount, vbInformation
            i = i + 1
        End If
        On Error GoTo 0
    Next net

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

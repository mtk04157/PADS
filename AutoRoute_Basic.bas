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

    On Error Resume Next
    For Each net In nets
        If net.PadCount >= 2 Then
            Call RouteNet(net)
            netCount = netCount + 1
        End If
    Next net
    On Error GoTo 0

    PCBDoc.Refresh

    MsgBox "Routed " & netCount & " nets", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub RouteNet(net)
    On Error Resume Next

    Dim pads
    Dim i
    Dim pad1, pad2
    Dim x1, y1, x2, y2

    Set pads = net.Pads

    If pads.Count >= 2 Then
        For i = 1 To pads.Count - 1
            Set pad1 = pads.Item(i)
            Set pad2 = pads.Item(i + 1)

            x1 = pad1.CenterX
            y1 = pad1.CenterY
            x2 = pad2.CenterX
            y2 = pad2.CenterY

            Call DrawTrace(net, x1, y1, x2, y2)
        Next i
    End If

    On Error GoTo 0
End Sub

Sub DrawTrace(net, x1, y1, x2, y2)
    On Error Resume Next

    If Abs(x1 - x2) > 0.001 Then
        Call DrawSegment(net, x1, y1, x2, y1)
    End If

    If Abs(y1 - y2) > 0.001 Then
        Call DrawSegment(net, x2, y1, x2, y2)
    End If

    On Error GoTo 0
End Sub

Sub DrawSegment(net, x1, y1, x2, y2)
    On Error Resume Next

    Dim segment
    Dim layer

    Set layer = ActiveDocument.LayerSet.Item("TOP")

    If Not (layer Is Nothing) Then
        Set segment = net.AddSegment()
        segment.X1 = x1
        segment.Y1 = y1
        segment.X2 = x2
        segment.Y2 = y2
        segment.Width = 0.254
        segment.Layer = layer
    End If

    On Error GoTo 0
End Sub

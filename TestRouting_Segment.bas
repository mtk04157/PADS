Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim nets
    Dim net
    Dim layer
    Dim segment
    Dim i

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set nets = PCBDoc.Nets
    Set layer = PCBDoc.LayerSet.Item("TOP")

    If layer Is Nothing Then
        MsgBox "TOP layer not found", vbCritical
        Exit Sub
    End If

    i = 0
    For Each net In nets
        On Error Resume Next
        If net.PadCount >= 2 And i = 0 Then
            Set segment = net.AddSegment()
            If Not (segment Is Nothing) Then
                segment.X1 = 10
                segment.Y1 = 10
                segment.X2 = 20
                segment.Y2 = 20
                segment.Width = 0.254
                segment.Layer = layer
                MsgBox "Created trace segment", vbInformation
            Else
                MsgBox "Failed to create segment", vbCritical
            End If
            i = i + 1
        End If
        On Error GoTo 0
    Next net

    PCBDoc.Refresh

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

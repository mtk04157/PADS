Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim i
    Dim info

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "No PCB file", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    i = 0

    For Each comp In components
        If i = 0 Then
            info = "Name: " & comp.Name
            On Error Resume Next
            info = info & " | PartNumber: " & comp.PartNumber
            info = info & " | RefDes: " & comp.RefDes
            On Error GoTo 0
            MsgBox info, vbInformation
        End If
        i = i + 1
    Next comp

    MsgBox "Step 2b complete", vbInformation
End Sub

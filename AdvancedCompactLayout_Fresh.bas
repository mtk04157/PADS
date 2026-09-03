Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim components
    Dim comp
    Dim compCount
    Dim i

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "Please open PCB file first", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    compCount = 0

    For Each comp In components
        If comp.selected Then
            compCount = compCount + 1
        End If
    Next comp

    If compCount = 0 Then
        MsgBox "Please select components", vbExclamation
        Exit Sub
    End If

    Call SmartCompactLayout(PCBDoc)
    Call AutoRoute(PCBDoc)

    MsgBox "Complete. Processed " & compCount & " components", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub SmartCompactLayout(PCBDoc)
    Dim components
    Dim comp
    Dim minX, minY
    Dim x, y
    Dim padSpacing
    Dim rowCount, colCount
    Dim row, col
    Dim i, compCount
    Dim width, height
    Dim isSpecial
    Dim specialCount
    Dim compName

    Set components = PCBDoc.Components
    padSpacing = 0.3
    specialCount = 0

    minX = 0
    minY = 0
    i = 0
    compCount = 0

    For Each comp In components
        If comp.selected Then
            If compCount = 0 Then
                minX = comp.LocationX
                minY = comp.LocationY
            End If
            compCount = compCount + 1
        End If
    Next comp

    If compCount = 0 Then Exit Sub

    For Each comp In components
        If comp.selected Then
            compName = UCase(comp.Name)
            If InStr(compName, "U6502") > 0 Or InStr(compName, "U6503") > 0 Then
                specialCount = specialCount + 1
            End If
        End If
    Next comp

    x = minX
    y = minY
    i = 0

    For Each comp In components
        If comp.selected Then
            compName = UCase(comp.Name)
            isSpecial = (InStr(compName, "U6502") > 0 Or InStr(compName, "U6503") > 0)

            If isSpecial Then
                comp.LocationX = x
                comp.LocationY = y

                width = Abs(comp.BBox.Xmax - comp.BBox.Xmin)
                x = x + width

                i = i + 1
            End If
        End If
    Next comp

    x = minX
    y = minY + 10
    rowCount = Int(Sqr(compCount - specialCount))
    If rowCount = 0 Then rowCount = 1
    colCount = Int((compCount - specialCount) / rowCount) + 1

    i = 0

    For Each comp In components
        If comp.selected Then
            compName = UCase(comp.Name)
            isSpecial = (InStr(compName, "U6502") > 0 Or InStr(compName, "U6503") > 0)

            If Not isSpecial Then
                row = i \ colCount
                col = i Mod colCount

                width = Abs(comp.BBox.Xmax - comp.BBox.Xmin) + 0.1
                If width < 0.01 Then width = 1

                height = Abs(comp.BBox.Ymax - comp.BBox.Ymin) + 0.1
                If height < 0.01 Then height = 1

                x = minX + (col * (width + padSpacing))
                y = minY + 10 + (row * (height + padSpacing))

                comp.LocationX = x
                comp.LocationY = y

                i = i + 1
            End If
        End If
    Next comp

    PCBDoc.Refresh
End Sub

Sub AutoRoute(PCBDoc)
    Dim nets
    Dim net

    Set nets = PCBDoc.Nets

    For Each net In nets
        If net.PadCount >= 2 Then
            Call RouteNet(net)
        End If
    Next net

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

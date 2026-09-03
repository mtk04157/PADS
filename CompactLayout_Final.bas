Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim components
    Dim comp
    Dim compCount, i
    Dim rowCount, colCount, row, col
    Dim x, y, nextX, nextY
    Dim width, height
    Dim minX, minY
    Dim padSpacing
    Dim isSpecial
    Dim compName
    Dim specialCount

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "Please open PCB file first", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    compCount = 0
    specialCount = 0
    padSpacing = 0.3

    For Each comp In components
        If comp.selected Then
            compCount = compCount + 1
            compName = UCase(comp.Name)
            If InStr(compName, "U6502") > 0 Or InStr(compName, "U6503") > 0 Then
                specialCount = specialCount + 1
            End If
        End If
    Next comp

    If compCount = 0 Then
        MsgBox "Please select components", vbExclamation
        Exit Sub
    End If

    nextX = 0
    nextY = 0
    i = 0

    For Each comp In components
        If comp.selected Then
            compName = UCase(comp.Name)
            isSpecial = (InStr(compName, "U6502") > 0 Or InStr(compName, "U6503") > 0)

            On Error Resume Next
            width = Abs(comp.BBox.Xmax - comp.BBox.Xmin) + padSpacing
            height = Abs(comp.BBox.Ymax - comp.BBox.Ymin) + padSpacing
            On Error GoTo 0

            If width < 0.1 Then width = 1
            If height < 0.1 Then height = 1

            If isSpecial Then
                comp.Move nextX, nextY
                nextX = nextX + width
            Else
                If i - specialCount > 0 Then
                    row = (i - specialCount) \ 4
                    col = (i - specialCount) Mod 4
                    x = col * width
                    y = nextY + (row * height) + 10
                    comp.Move x, y
                End If
            End If

            i = i + 1
        End If
    Next comp

    On Error Resume Next
    PCBDoc.Refresh
    On Error GoTo 0

    MsgBox "Layout complete. Processed " & compCount & " components", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

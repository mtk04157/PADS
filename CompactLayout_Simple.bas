Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim components
    Dim comp
    Dim compCount, i
    Dim row, col
    Dim deltaX, deltaY
    Dim spacing
    Dim compName

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "Please open PCB file first", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components
    compCount = 0
    spacing = 3

    For Each comp In components
        If comp.selected Then
            compCount = compCount + 1
        End If
    Next comp

    If compCount = 0 Then
        MsgBox "Please select components", vbExclamation
        Exit Sub
    End If

    i = 0
    For Each comp In components
        If comp.selected Then
            compName = UCase(comp.Name)

            If InStr(compName, "U6502") > 0 Or InStr(compName, "U6503") > 0 Then
                deltaX = i * spacing
                deltaY = 0
            Else
                row = (i - 0) \ 4
                col = (i - 0) Mod 4
                deltaX = col * spacing - (i * spacing)
                deltaY = row * spacing
            End If

            comp.Move deltaX, deltaY
            i = i + 1
        End If
    Next comp

    On Error Resume Next
    PCBDoc.Refresh
    On Error GoTo 0

    MsgBox "Arranged " & compCount & " components", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

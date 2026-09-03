Sub Main
    Dim PCBDoc
    Dim components
    Dim comp
    Dim compCount

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "PCB file not open"
        Exit Sub
    End If

    Set components = PCBDoc.Components
    compCount = 0

    '计数选中的组件
    For Each comp In components
        If comp.selected Then
            compCount = compCount + 1
        End If
    Next comp

    MsgBox "Found " & compCount & " selected components"
End Sub

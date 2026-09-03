'================================================
' PADS VX2.13 脚本: 智能紧凑布局
'================================================

Const MIN_PAD_SPACING = 0.2
Const MIN_OUTLINE_SPACING = 0
Const TRACE_WIDTH = 0.254

Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim components
    Dim comp
    Dim compCount

    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "请先打开PADS PCB文件", vbCritical
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

    If compCount = 0 Then
        MsgBox "请先选择要布局的组件区块", vbExclamation
        Exit Sub
    End If

    '执行智能布局
    Call SmartCompactLayout(PCBDoc)

    '执行走线
    Call OptimizedAutoRoute(PCBDoc)

    MsgBox "智能紧凑布局和走线完成!" & vbCrLf & _
           "已处理 " & compCount & " 个组件", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
End Sub

'智能紧凑布局
Sub SmartCompactLayout(PCBDoc)
    Dim components
    Dim comp
    Dim specialCount, regularCount
    Dim minX, minY
    Dim x, y
    Dim padSpacing
    Dim i, row, col, rowCount, colCount
    Dim width, height
    Dim isSpecial
    Dim lastSpecialX, lastSpecialY, lastSpecialWidth
    Dim compName

    Set components = PCBDoc.Components
    padSpacing = MIN_PAD_SPACING + 0.1

    '先处理特殊组件
    lastSpecialX = 0
    lastSpecialY = 0
    lastSpecialWidth = 0

    For Each comp In components
        If comp.selected Then
            compName = UCase(comp.Name)
            isSpecial = False

            If InStr(compName, "U6502") > 0 Or InStr(compName, "U6503") > 0 Then
                isSpecial = True
            End If

            If isSpecial Then
                width = Abs(comp.BBox.Xmax - comp.BBox.Xmin) + 0.1
                If width < 0.01 Then width = 1

                If lastSpecialWidth = 0 Then
                    comp.LocationX = 0
                    comp.LocationY = 0
                    lastSpecialX = 0
                    lastSpecialY = 0
                    lastSpecialWidth = width
                Else
                    x = lastSpecialX + lastSpecialWidth + MIN_OUTLINE_SPACING
                    comp.LocationX = x
                    comp.LocationY = lastSpecialY
                    lastSpecialX = x
                    lastSpecialWidth = width
                End If
            End If
        End If
    Next comp

    '然后处理普通组件
    minX = lastSpecialX + lastSpecialWidth + padSpacing
    If minX < 0 Then minX = 0
    minY = 0
    regularCount = 0

    For Each comp In components
        If comp.selected Then
            compName = UCase(comp.Name)
            isSpecial = False

            If InStr(compName, "U6502") > 0 Or InStr(compName, "U6503") > 0 Then
                isSpecial = True
            End If

            If Not isSpecial Then
                regularCount = regularCount + 1
            End If
        End If
    Next comp

    If regularCount > 0 Then
        rowCount = Int(Sqr(regularCount))
        If rowCount = 0 Then rowCount = 1
        colCount = Int(regularCount / rowCount) + 1

        i = 0
        x = minX
        y = minY

        For Each comp In components
            If comp.selected Then
                compName = UCase(comp.Name)
                isSpecial = False

                If InStr(compName, "U6502") > 0 Or InStr(compName, "U6503") > 0 Then
                    isSpecial = True
                End If

                If Not isSpecial Then
                    width = Abs(comp.BBox.Xmax - comp.BBox.Xmin) + 0.1
                    If width < 0.01 Then width = 1

                    height = Abs(comp.BBox.Ymax - comp.BBox.Ymin) + 0.1
                    If height < 0.01 Then height = 1

                    comp.LocationX = x
                    comp.LocationY = y

                    x = x + width + padSpacing

                    If (i + 1) Mod colCount = 0 Then
                        x = minX
                        y = y + height + padSpacing
                    End If

                    i = i + 1
                End If
            End If
        Next comp
    End If

    PCBDoc.Refresh
End Sub

'优化自动走线
Sub OptimizedAutoRoute(PCBDoc)
    Dim nets
    Dim net

    Set nets = PCBDoc.Nets

    For Each net In nets
        If net.PadCount >= 2 Then
            Call OptimizeNetRoute(net)
        End If
    Next net

End Sub

'单网络最优走线
Sub OptimizeNetRoute(net)
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

            Call DrawStraightTrace(net, x1, y1, x2, y2)
        Next i
    End If

    On Error GoTo 0
End Sub

'绘制直线走线
Sub DrawStraightTrace(net, x1, y1, x2, y2)
    On Error Resume Next

    '先水平后竖直
    If Abs(x1 - x2) > 0.001 Then
        Call AddTraceSegment(net, x1, y1, x2, y1)
    End If

    If Abs(y1 - y2) > 0.001 Then
        Call AddTraceSegment(net, x2, y1, x2, y2)
    End If

    On Error GoTo 0
End Sub

'添加走线段
Sub AddTraceSegment(net, x1, y1, x2, y2)
    On Error Resume Next

    Dim seg
    Dim layer

    Set layer = ActiveDocument.LayerSet.Item("TOP")

    If Not (layer Is Nothing) Then
        Set seg = net.AddSegment()
        seg.X1 = x1
        seg.Y1 = y1
        seg.X2 = x2
        seg.Y2 = y2
        seg.Width = TRACE_WIDTH
        seg.Layer = layer
    End If

    On Error GoTo 0
End Sub

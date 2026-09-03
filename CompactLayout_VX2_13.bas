'================================================
' PADS VX2.13 脚本: 紧凑布局和直线走线
' 功能: 选定区块内的元件以0.2mm最小间距摆放
'       并进行最短距离的直线走线
'================================================

Const MIN_PAD_SPACING = 0.2
Const TRACE_WIDTH = 0.254

Sub Main
    On Error GoTo ErrorHandler

    Dim PCBDoc
    Dim components
    Dim comp
    Dim compCount
    Dim i

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
        MsgBox "请先选择要布局的组件", vbExclamation
        Exit Sub
    End If

    '执行紧凑布局
    Call CompactLayout(PCBDoc)

    '执行直线走线
    Call AutoRoute(PCBDoc)

    MsgBox "布局和走线完成!" & vbCrLf & _
           "已处理 " & compCount & " 个组件", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
End Sub

'紧凑布局函数
Sub CompactLayout(PCBDoc)
    Dim components
    Dim comp
    Dim minX, minY
    Dim x, y
    Dim padSpacing
    Dim rowCount, colCount
    Dim row, col
    Dim i, compCount
    Dim width, height

    Set components = PCBDoc.Components
    padSpacing = MIN_PAD_SPACING + 0.1

    minX = 0
    minY = 0
    i = 0
    compCount = 0

    '第一遍：计数和找起点
    For Each comp In components
        If comp.selected Then
            If compCount = 0 Then
                minX = comp.LocationX
                minY = comp.LocationY
            End If
            compCount = compCount + 1
        End If
    Next comp

    '计算网格
    rowCount = Int(Sqr(compCount))
    If rowCount = 0 Then rowCount = 1
    colCount = Int(compCount / rowCount) + 1

    '第二遍：摆放组件
    i = 0
    For Each comp In components
        If comp.selected Then
            row = i \ colCount
            col = i Mod colCount

            width = Abs(comp.BBox.Xmax - comp.BBox.Xmin) + 0.1
            If width < 0.01 Then width = 1

            height = Abs(comp.BBox.Ymax - comp.BBox.Ymin) + 0.1
            If height < 0.01 Then height = 1

            x = minX + (col * (width + padSpacing))
            y = minY + (row * (height + padSpacing))

            comp.LocationX = x
            comp.LocationY = y

            i = i + 1
        End If
    Next comp

    PCBDoc.Refresh
End Sub

'自动走线函数
Sub AutoRoute(PCBDoc)
    Dim nets
    Dim net
    Dim routedCount

    Set nets = PCBDoc.Nets
    routedCount = 0

    For Each net In nets
        If net.PadCount >= 2 Then
            If RouteNet(net) Then
                routedCount = routedCount + 1
            End If
        End If
    Next net

End Sub

'单个网络走线
Function RouteNet(net)
    On Error GoTo ErrorHandler

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

    RouteNet = True
    Exit Function

ErrorHandler:
    RouteNet = False
End Function

'绘制走线
Sub DrawTrace(net, x1, y1, x2, y2)
    On Error Resume Next

    '先水平后竖直 (L形走线)
    If Abs(x1 - x2) > 0.001 Then
        Call DrawSegment(net, x1, y1, x2, y1)
    End If

    If Abs(y1 - y2) > 0.001 Then
        Call DrawSegment(net, x2, y1, x2, y2)
    End If

    On Error GoTo 0
End Sub

'绘制单个走线段
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
        segment.Width = TRACE_WIDTH
        segment.Layer = layer
    End If

    On Error GoTo 0
End Sub

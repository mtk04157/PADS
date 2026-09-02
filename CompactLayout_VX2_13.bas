'================================================
' PADS VX2.13 脚本: 紧凑布局和直线走线
' 功能: 选定区块内的元件以0.2mm最小间距摆放
'       并进行最短距离的直线走线
'================================================

Sub Main()
    On Error GoTo ErrorHandler

    Dim PCBDoc As Object
    Dim components As Object
    Dim selectedComps As New Collection
    Dim comp As Object
    Dim i As Integer

    '获取当前PCB文档
    Set PCBDoc = ActiveDocument

    If PCBDoc Is Nothing Then
        MsgBox "请先打开PADS PCB文件", vbCritical
        Exit Sub
    End If

    Set components = PCBDoc.Components

    '收集选中的组件
    For Each comp In components
        If comp.selected Then
            selectedComps.Add comp
        End If
    Next comp

    If selectedComps.Count = 0 Then
        MsgBox "请先选择要布局的组件", vbExclamation
        Exit Sub
    End If

    '执行紧凑布局
    Call CompactLayout(PCBDoc, selectedComps)

    '执行直线走线
    Call AutoRoute(PCBDoc)

    MsgBox "布局和走线完成!" & vbCrLf & _
           "已处理 " & selectedComps.Count & " 个组件", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
End Sub

'紧凑布局函数 - 按照0.2mm最小间距摆放
Sub CompactLayout(PCBDoc As Object, selectedComps As Collection)
    Dim minX As Double, minY As Double
    Dim padSpacing As Double
    Dim rowCount As Integer, colCount As Integer
    Dim x As Double, y As Double
    Dim row As Integer, col As Integer
    Dim i As Integer
    Dim comp As Object

    '常数定义
    Const MIN_PAD_SPACING = 0.2
    const COMPONENT_SPACING = 0.1

    padSpacing = MIN_PAD_SPACING + COMPONENT_SPACING

    '设置基准点
    Set comp = selectedComps.Item(1)
    minX = comp.LocationX
    minY = comp.LocationY

    '计算网格排列
    rowCount = Int(Sqr(selectedComps.Count))
    If rowCount = 0 Then rowCount = 1
    colCount = Int(selectedComps.Count / rowCount) + 1

    '摆放组件
    For i = 1 To selectedComps.Count
        Set comp = selectedComps.Item(i)

        '计算行列位置
        row = (i - 1) \ colCount
        col = (i - 1) Mod colCount

        '计算新位置
        x = minX + (col * (GetComponentWidth(comp) + padSpacing))
        y = minY + (row * (GetComponentHeight(comp) + padSpacing))

        '移动组件
        comp.LocationX = x
        comp.LocationY = y

        Debug.Print "组件 " & comp.Name & " 已移到: (" & Format(x, "0.00") & ", " & Format(y, "0.00") & ")"
    Next i

    PCBDoc.Refresh
End Sub

'自动走线函数
Sub AutoRoute(PCBDoc As Object)
    Dim nets As Object
    Dim net As Object
    Dim i As Integer
    Dim routedCount As Integer

    Set nets = PCBDoc.Nets
    routedCount = 0

    '对每个网络进行走线
    For Each net In nets
        If net.PadCount >= 2 Then
            If RouteNet(PCBDoc, net) Then
                routedCount = routedCount + 1
            End If
        End If
    Next net

    Debug.Print "已完成 " & routedCount & " 个网络的走线"
    PCBDoc.Refresh
End Sub

'单个网络走线
Function RouteNet(PCBDoc As Object, net As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim pads As Object
    Dim i As Integer
    Dim pad1 As Object, pad2 As Object
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double

    Set pads = net.Pads

    '连接所有PAD
    If pads.Count >= 2 Then
        For i = 1 To pads.Count - 1
            Set pad1 = pads.Item(i)
            Set pad2 = pads.Item(i + 1)

            x1 = pad1.CenterX
            y1 = pad1.CenterY
            x2 = pad2.CenterX
            y2 = pad2.CenterY

            '绘制L形走线
            Call DrawTrace(net, x1, y1, x2, y2)
        Next i
    End If

    RouteNet = True
    Exit Function

ErrorHandler:
    RouteNet = False
End Function

'绘制走线
Sub DrawTrace(net As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    On Error Resume Next

    Const TRACE_WIDTH = 0.254

    '先水平后竖直 (L形走线)
    If Abs(x1 - x2) > 0.001 Then
        Call DrawSegment(net, x1, y1, x2, y1, TRACE_WIDTH)
    End If

    If Abs(y1 - y2) > 0.001 Then
        Call DrawSegment(net, x2, y1, x2, y2, TRACE_WIDTH)
    End If

    On Error GoTo 0
End Sub

'绘制单个走线段
Sub DrawSegment(net As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double, width As Double)
    On Error Resume Next

    Dim segment As Object
    Dim layer As Object

    '添加走线段到网络
    Set segment = net.AddSegment()
    segment.X1 = x1
    segment.Y1 = y1
    segment.X2 = x2
    segment.Y2 = y2
    segment.Width = width
    segment.Layer = ActiveDocument.LayerSet.Item("TOP")

    On Error GoTo 0
End Sub

'获取组件宽度
Function GetComponentWidth(comp As Object) As Double
    On Error Resume Next
    Dim width As Double
    width = Abs(comp.BBox.Xmax - comp.BBox.Xmin)
    If width < 0.01 Then width = 1
    GetComponentWidth = width
    On Error GoTo 0
End Function

'获取组件高度
Function GetComponentHeight(comp As Object) As Double
    On Error Resume Next
    Dim height As Double
    height = Abs(comp.BBox.Ymax - comp.BBox.Ymin)
    If height < 0.01 Then height = 1
    GetComponentHeight = height
    On Error GoTo 0
End Function

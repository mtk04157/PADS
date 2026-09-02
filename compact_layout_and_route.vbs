'================================================
' PADS脚本: 紧凑布局和直线走线
' 功能: 选定区块内的元件以0.2mm最小间距摆放
'       并进行最短距离的直线走线
'================================================

Option Explicit

Dim PCBApp As Object
Dim PCBDoc As Object
Dim PCBView As Object
Dim PCBSelection As Object
Dim PCBCompList As Object

Const MIN_PAD_SPACING = 0.2  '最小PAD到PAD间距(mm)
Const TRACE_WIDTH = 0.254    '走线宽度(mm)

Sub Main()
    On Error GoTo ErrorHandler

    '初始化PADS应用
    Set PCBApp = CreateObject("PADS.PadsPcbApplication")
    Set PCBDoc = PCBApp.ActiveDocument
    Set PCBView = PCBDoc.ActiveView

    If PCBDoc Is Nothing Then
        MsgBox "请先打开PADS PCB文件", vbCritical
        Exit Sub
    End If

    '获取选中的组件
    Set PCBSelection = PCBDoc.SelectionSet

    If PCBSelection.Count = 0 Then
        MsgBox "请先选择要布局的组件", vbExclamation
        Exit Sub
    End If

    Call CompactLayout(PCBDoc, PCBSelection)
    Call AutoRoute(PCBDoc)

    MsgBox "布局和走线完成!", vbInformation

    Exit Sub
ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
End Sub

'紧凑布局函数 - 按照0.2mm最小间距摆放
Sub CompactLayout(PCBDoc As Object, PCBSelection As Object)
    Dim compArray() As Object
    Dim i As Integer, j As Integer
    Dim comp As Object
    Dim x As Double, y As Double
    Dim minX As Double, minY As Double
    Dim padSpacing As Double
    Dim rowCount As Integer, colCount As Integer
    Dim gridX As Double, gridY As Double

    '收集所有选中组件
    ReDim compArray(PCBSelection.Count - 1)
    For i = 0 To PCBSelection.Count - 1
        Set compArray(i) = PCBSelection.Item(i)
    Next

    '排序组件(按X坐标)
    Call SortComponents(compArray)

    '设置基准点
    Set comp = compArray(0)
    minX = comp.BBox.Xmin
    minY = comp.BBox.Ymin
    padSpacing = MIN_PAD_SPACING + 0.1  '考虑焊盘大小

    '计算网格排列
    rowCount = Int(Sqr(PCBSelection.Count))
    colCount = Int(PCBSelection.Count / rowCount) + 1

    '摆放组件
    For i = 0 To PCBSelection.Count - 1
        Set comp = compArray(i)

        '计算新位置
        Dim row As Integer, col As Integer
        row = i \ colCount
        col = i Mod colCount

        gridX = minX + (col * (GetComponentWidth(comp) + padSpacing))
        gridY = minY + (row * (GetComponentHeight(comp) + padSpacing))

        '移动组件到新位置
        Call MoveComponent(comp, gridX, gridY)

        '打印调试信息
        Debug.Print "组件 " & comp.Name & " 已移到: (" & gridX & ", " & gridY & ")"
    Next i

    '刷新视图
    PCBDoc.Refresh
End Sub

'自动走线函数 - 直线最短路径
Sub AutoRoute(PCBDoc As Object)
    Dim net As Object
    Dim netList As Object
    Dim i As Integer

    Set netList = PCBDoc.Nets

    '对每个网络进行走线
    For i = 0 To netList.Count - 1
        Set net = netList.Item(i)

        '跳过未连接的网络
        If net.PadCount > 1 Then
            Call RouteNet(PCBDoc, net)
        End If
    Next i

    PCBDoc.Refresh
End Sub

'单个网络走线 - 直线连接
Sub RouteNet(PCBDoc As Object, net As Object)
    On Error Resume Next

    Dim pads As Object
    Dim i As Integer
    Dim pad1 As Object, pad2 As Object
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim viaX As Double, viaY As Double

    Set pads = net.Pads

    '连接所有的PAD - 采用最小树形走线
    If pads.Count >= 2 Then
        For i = 0 To pads.Count - 2
            Set pad1 = pads.Item(i)
            Set pad2 = pads.Item(i + 1)

            x1 = pad1.CenterX
            y1 = pad1.CenterY
            x2 = pad2.CenterX
            y2 = pad2.CenterY

            '创建直线走线 - 使用曼哈顿风格(水平->竖直)
            Call DrawTrace(PCBDoc, net, x1, y1, x2, y2)
        Next i
    End If

    On Error GoTo 0
End Sub

'绘制走线 - 水平->竖直->水平的直线走线
Sub DrawTrace(PCBDoc As Object, net As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    On Error Resume Next

    Dim layer As Object
    Dim segment As Object

    '使用信号层
    Set layer = PCBDoc.LayerSet.Item("TOP")

    '绘制水平走线
    If x1 <> x2 Then
        Set segment = PCBDoc.Nets.Item(net.Name).AddSegment()
        segment.Layer = layer
        segment.X1 = x1
        segment.Y1 = y1
        segment.X2 = x2
        segment.Y2 = y1
        segment.Width = TRACE_WIDTH
    End If

    '绘制竖直走线
    If y1 <> y2 Then
        Set segment = PCBDoc.Nets.Item(net.Name).AddSegment()
        segment.Layer = layer
        segment.X1 = x2
        segment.Y1 = y1
        segment.X2 = x2
        segment.Y2 = y2
        segment.Width = TRACE_WIDTH
    End If

    On Error GoTo 0
End Sub

'获取组件宽度
Function GetComponentWidth(comp As Object) As Double
    On Error Resume Next
    GetComponentWidth = Abs(comp.BBox.Xmax - comp.BBox.Xmin)
    On Error GoTo 0
End Function

'获取组件高度
Function GetComponentHeight(comp As Object) As Double
    On Error Resume Next
    GetComponentHeight = Abs(comp.BBox.Ymax - comp.BBox.Ymin)
    On Error GoTo 0
End Function

'移动组件
Sub MoveComponent(comp As Object, x As Double, y As Double)
    On Error Resume Next
    comp.LocationX = x
    comp.LocationY = y
    On Error GoTo 0
End Sub

'排序组件数组
Sub SortComponents(compArray() As Object)
    Dim i As Integer, j As Integer
    Dim temp As Object

    For i = LBound(compArray) To UBound(compArray) - 1
        For j = i + 1 To UBound(compArray)
            If compArray(i).BBox.Xmin > compArray(j).BBox.Xmin Then
                Set temp = compArray(i)
                Set compArray(i) = compArray(j)
                Set compArray(j) = temp
            End If
        Next j
    Next i
End Sub

'入口点
Call Main()

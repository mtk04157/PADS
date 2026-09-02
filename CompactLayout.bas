'================================================
' PADS VX2.13 BASIC脚本: 紧凑布局和直线走线
' 功能: 选定区块内的元件以0.2mm最小间距摆放
'       并进行最短距离的直线走线
'================================================

Option Explicit

Dim PCBApp As Object
Dim PCBDoc As Object
Dim PCBView As Object
Dim PCBSelection As Object

Const MIN_PAD_SPACING = 0.2      '最小PAD到PAD间距(mm)
Const TRACE_WIDTH = 0.254        '走线宽度(mm)

Sub Main()
    On Error GoTo ErrorHandler

    '初始化PADS应用
    Set PCBApp = CreateObject("PADS.PadsPcbApplication")
    If PCBApp Is Nothing Then
        MsgBox "无法启动PADS应用", vbCritical
        Exit Sub
    End If

    Set PCBDoc = PCBApp.ActiveDocument
    If PCBDoc Is Nothing Then
        MsgBox "请先打开PADS PCB文件", vbCritical
        Exit Sub
    End If

    Set PCBView = PCBDoc.ActiveView
    Set PCBSelection = PCBDoc.SelectionSet

    If PCBSelection.Count = 0 Then
        MsgBox "请先选择要布局的组件", vbExclamation
        Exit Sub
    End If

    '执行紧凑布局
    Call CompactLayout(PCBDoc, PCBSelection)

    '执行直线走线
    Call AutoRoute(PCBDoc)

    MsgBox "布局和走线完成!" & vbCrLf & _
           "已处理 " & PCBSelection.Count & " 个组件", vbInformation

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
    Dim row As Integer, col As Integer

    '收集所有选中组件
    ReDim compArray(PCBSelection.Count - 1)
    For i = 0 To PCBSelection.Count - 1
        Set compArray(i) = PCBSelection.Item(i)
    Next

    '排序组件(按X坐标)
    Call SortComponents(compArray)

    '设置基准点
    Set comp = compArray(0)
    minX = comp.LocationX
    minY = comp.LocationY
    padSpacing = MIN_PAD_SPACING + 0.1

    '计算网格排列
    rowCount = Int(Sqr(PCBSelection.Count))
    If rowCount = 0 Then rowCount = 1
    colCount = Int(PCBSelection.Count / rowCount) + 1

    '摆放组件
    For i = 0 To PCBSelection.Count - 1
        Set comp = compArray(i)

        '计算新位置
        row = i \ colCount
        col = i Mod colCount

        x = minX + (col * (GetComponentWidth(comp) + padSpacing))
        y = minY + (row * (GetComponentHeight(comp) + padSpacing))

        '移动组件到新位置
        comp.LocationX = x
        comp.LocationY = y

        Debug.Print "组件 " & comp.Name & " 已移到: (" & Format(x, "0.00") & ", " & Format(y, "0.00") & ")"
    Next i

    '刷新视图
    PCBDoc.Refresh
End Sub

'自动走线函数 - 直线最短路径
Sub AutoRoute(PCBDoc As Object)
    Dim netList As Object
    Dim i As Integer
    Dim net As Object
    Dim routedCount As Integer

    Set netList = PCBDoc.Nets
    routedCount = 0

    '对每个网络进行走线
    For i = 0 To netList.Count - 1
        Set net = netList.Item(i)

        If net.PadCount >= 2 Then
            If RouteNet(PCBDoc, net) Then
                routedCount = routedCount + 1
            End If
        End If
    Next i

    Debug.Print "已完成 " & routedCount & " 个网络的走线"
    PCBDoc.Refresh
End Sub

'单个网络走线 - 直线连接
Function RouteNet(PCBDoc As Object, net As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim pads As Object
    Dim i As Integer
    Dim pad1 As Object, pad2 As Object
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double

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

    RouteNet = True
    Exit Function

ErrorHandler:
    RouteNet = False
End Function

'绘制走线 - 水平->竖直的直线走线
Sub DrawTrace(PCBDoc As Object, net As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    On Error Resume Next

    '绘制水平走线
    If Abs(x1 - x2) > 0.001 Then
        Call DrawSegment(PCBDoc, net, x1, y1, x2, y1)
    End If

    '绘制竖直走线
    If Abs(y1 - y2) > 0.001 Then
        Call DrawSegment(PCBDoc, net, x2, y1, x2, y2)
    End If

    On Error GoTo 0
End Sub

'绘制单个走线段
Sub DrawSegment(PCBDoc As Object, net As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    On Error Resume Next

    Dim segment As Object
    Dim layer As Object

    '获取顶层
    Set layer = PCBDoc.LayerSet.Item("TOP")

    If Not layer Is Nothing Then
        Set segment = net.AddSegment()
        segment.Layer = layer
        segment.X1 = x1
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
    Dim width As Double
    width = Abs(comp.BBox.Xmax - comp.BBox.Xmin)
    If width < 0.01 Then width = 1  '默认值
    GetComponentWidth = width
    On Error GoTo 0
End Function

'获取组件高度
Function GetComponentHeight(comp As Object) As Double
    On Error Resume Next
    Dim height As Double
    height = Abs(comp.BBox.Ymax - comp.BBox.Ymin)
    If height < 0.01 Then height = 1  '默认值
    GetComponentHeight = height
    On Error GoTo 0
End Function

'排序组件数组
Sub SortComponents(compArray() As Object)
    Dim i As Integer, j As Integer
    Dim temp As Object
    Dim len As Integer

    len = UBound(compArray)

    For i = LBound(compArray) To len - 1
        For j = i + 1 To len
            On Error Resume Next
            If compArray(i).LocationX > compArray(j).LocationX Then
                Set temp = compArray(i)
                Set compArray(i) = compArray(j)
                Set compArray(j) = temp
            End If
            On Error GoTo 0
        Next j
    Next i
End Sub

'入口点
Call Main()

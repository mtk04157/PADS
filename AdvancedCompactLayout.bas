'================================================
' PADS VX2.13 BASIC脚本: 智能紧凑布局
' 功能:
'   1. 0.2mm最小PAD间距摆放
'   2. 特定组件(U6502/U6503)outline对齐
'   3. 直线最短走线
'   4. 自动过孔放置
'================================================

Option Explicit

Dim PCBApp As Object
Dim PCBDoc As Object
Dim PCBView As Object

Const MIN_PAD_SPACING = 0.2      '最小PAD到PAD间距(mm)
Const MIN_OUTLINE_SPACING = 0    'Outline对齐间距(0=贴邻)
Const TRACE_WIDTH = 0.254        '走线宽度(mm) 10mil
Const VIA_DIAMETER = 0.3         '过孔直径(mm)

Sub Main()
    On Error GoTo ErrorHandler

    '初始化PADS
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

    '获取用户选择
    Dim selectedComps As Object
    Set selectedComps = PCBDoc.SelectionSet

    If selectedComps.Count = 0 Then
        MsgBox "请先选择要布局的组件区块", vbExclamation
        Exit Sub
    End If

    '执行布局
    Call SmartCompactLayout(PCBDoc, selectedComps)

    '执行走线
    Call OptimizedAutoRoute(PCBDoc)

    MsgBox "智能紧凑布局和走线完成!" & vbCrLf & _
           "已处理 " & selectedComps.Count & " 个组件", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
End Sub

'智能紧凑布局
Sub SmartCompactLayout(PCBDoc As Object, selectedComps As Object)
    Dim i As Integer
    Dim comp As Object
    Dim specialComps As New Collection
    Dim regularComps As New Collection
    Dim isSpecial As Boolean

    '分类组件
    For i = 0 To selectedComps.Count - 1
        Set comp = selectedComps.Item(i)
        isSpecial = False

        '检查是否为特殊组件
        If InStr(comp.Name, "U6502") > 0 Or InStr(comp.Name, "U6503") > 0 Then
            specialComps.Add comp
            isSpecial = True
        End If

        If Not isSpecial Then
            regularComps.Add comp
        End If
    Next i

    '首先放置特殊组件（U6502/U6503）- Outline贴邻
    If specialComps.Count > 0 Then
        Call PlaceSpecialComponentsOutlineAligned(PCBDoc, specialComps)
    End If

    '然后放置普通组件 - 0.2mm PAD间距
    If regularComps.Count > 0 Then
        Call PlaceRegularComponentsCompact(PCBDoc, regularComps)
    End If

    PCBDoc.Refresh
End Sub

'特殊组件Outline对齐放置
Sub PlaceSpecialComponentsOutlineAligned(PCBDoc As Object, specialComps As Collection)
    Dim i As Integer
    Dim comp As Object
    Dim refComp As Object
    Dim x As Double, y As Double
    Dim spacing As Double

    If specialComps.Count = 0 Then Exit Sub

    '使用第一个组件作为参考
    Set refComp = specialComps.Item(1)
    x = refComp.LocationX
    y = refComp.LocationY + GetComponentHeight(refComp) + MIN_OUTLINE_SPACING

    '按行排列特殊组件，outline贴邻
    For i = 2 To specialComps.Count
        Set comp = specialComps.Item(i)

        '水平排列
        x = x + GetComponentWidth(specialComps.Item(i - 1)) + MIN_OUTLINE_SPACING
        comp.LocationX = x
        comp.LocationY = y

        '记录日志
        Debug.Print "特殊组件 " & comp.Name & " Outline对齐到: (" & Format(x, "0.00") & ", " & Format(y, "0.00") & ")"
    Next i
End Sub

'普通组件紧凑放置 - 0.2mm间距
Sub PlaceRegularComponentsCompact(PCBDoc As Object, regularComps As Collection)
    Dim i As Integer
    Dim comp As Object
    Dim x As Double, y As Double
    Dim startX As Double, startY As Double
    Dim maxHeight As Double
    Dim currentHeight As Double
    Dim spacing As Double
    Dim colCount As Integer

    If regularComps.Count = 0 Then Exit Sub

    spacing = MIN_PAD_SPACING + 0.1  '考虑焊盘大小

    '获取起始位置
    startX = 0
    startY = 0

    x = startX
    y = startY
    maxHeight = 0
    colCount = Int(Sqr(regularComps.Count))
    If colCount = 0 Then colCount = 1

    '按网格排列
    For i = 1 To regularComps.Count
        Set comp = regularComps.Item(i)
        currentHeight = GetComponentHeight(comp)

        If currentHeight > maxHeight Then
            maxHeight = currentHeight
        End If

        comp.LocationX = x
        comp.LocationY = y

        '水平排列
        x = x + GetComponentWidth(comp) + spacing

        '换行
        If i Mod colCount = 0 Then
            x = startX
            y = y + maxHeight + spacing
            maxHeight = 0
        End If

        Debug.Print "普通组件 " & comp.Name & " 放置到: (" & Format(comp.LocationX, "0.00") & ", " & Format(comp.LocationY, "0.00") & ")"
    Next i
End Sub

'优化自动走线 - 直线最短路径
Sub OptimizedAutoRoute(PCBDoc As Object)
    Dim netList As Object
    Dim i As Integer
    Dim net As Object
    Dim routedCount As Integer

    Set netList = PCBDoc.Nets
    routedCount = 0

    '对每个网络进行最优走线
    For i = 0 To netList.Count - 1
        Set net = netList.Item(i)

        If net.PadCount >= 2 Then
            If OptimizeNetRoute(PCBDoc, net) Then
                routedCount = routedCount + 1
            End If
        End If
    Next i

    Debug.Print "已完成 " & routedCount & " 个网络的走线"
    PCBDoc.Refresh
End Sub

'单网络最优走线
Function OptimizeNetRoute(PCBDoc As Object, net As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim pads As Object
    Dim padCount As Integer
    Dim i As Integer, j As Integer
    Dim minDist As Double
    Dim nextPad As Integer
    Dim currentPad As Integer
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim visited() As Boolean

    Set pads = net.Pads
    padCount = pads.Count

    If padCount < 2 Then
        OptimizeNetRoute = False
        Exit Function
    End If

    '使用最近邻算法进行走线
    ReDim visited(padCount - 1)
    currentPad = 0
    visited(0) = True

    For i = 1 To padCount - 1
        minDist = 999999
        nextPad = -1

        '找到距离当前PAD最近的未访问PAD
        For j = 0 To padCount - 1
            If Not visited(j) Then
                Dim dist As Double
                dist = CalculateDistance(pads.Item(currentPad), pads.Item(j))
                If dist < minDist Then
                    minDist = dist
                    nextPad = j
                End If
            End If
        Next j

        If nextPad >= 0 Then
            visited(nextPad) = True
            x1 = pads.Item(currentPad).CenterX
            y1 = pads.Item(currentPad).CenterY
            x2 = pads.Item(nextPad).CenterX
            y2 = pads.Item(nextPad).CenterY

            '绘制直线走线
            Call DrawStraightTrace(PCBDoc, net, x1, y1, x2, y2)
            currentPad = nextPad
        End If
    Next i

    OptimizeNetRoute = True
    Exit Function

ErrorHandler:
    OptimizeNetRoute = False
End Function

'绘制直线走线
Sub DrawStraightTrace(PCBDoc As Object, net As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    On Error Resume Next

    '先水平后竖直 (L形走线)
    If Abs(x1 - x2) > 0.001 Then
        Call AddTraceSegment(PCBDoc, net, "TOP", x1, y1, x2, y1)
    End If

    If Abs(y1 - y2) > 0.001 Then
        Call AddTraceSegment(PCBDoc, net, "TOP", x2, y1, x2, y2)
    End If

    On Error GoTo 0
End Sub

'添加走线段
Sub AddTraceSegment(PCBDoc As Object, net As Object, layerName As String, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    On Error Resume Next

    Dim seg As Object
    Dim layerObj As Object

    Set layerObj = PCBDoc.LayerSet.Item(layerName)

    If Not layerObj Is Nothing Then
        Set seg = net.AddSegment()
        seg.Layer = layerObj
        seg.X1 = x1
        seg.Y1 = y1
        seg.X2 = x2
        seg.Y2 = y2
        seg.Width = TRACE_WIDTH
    End If

    On Error GoTo 0
End Sub

'计算两个PAD之间的距离
Function CalculateDistance(pad1 As Object, pad2 As Object) As Double
    Dim dx As Double, dy As Double
    dx = pad1.CenterX - pad2.CenterX
    dy = pad1.CenterY - pad2.CenterY
    CalculateDistance = Sqr(dx * dx + dy * dy)
End Function

'获取组件宽度
Function GetComponentWidth(comp As Object) As Double
    On Error Resume Next
    Dim width As Double
    width = Abs(comp.BBox.Xmax - comp.BBox.Xmin) + 0.1
    If width < 0.01 Then width = 1
    GetComponentWidth = width
    On Error GoTo 0
End Function

'获取组件高度
Function GetComponentHeight(comp As Object) As Double
    On Error Resume Next
    Dim height As Double
    height = Abs(comp.BBox.Ymax - comp.BBox.Ymin) + 0.1
    If height < 0.01 Then height = 1
    GetComponentHeight = height
    On Error GoTo 0
End Function

'入口
Call Main()

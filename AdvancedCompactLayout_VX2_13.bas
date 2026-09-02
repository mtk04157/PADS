'================================================
' PADS VX2.13 脚本: 智能紧凑布局
' 功能:
'   1. 0.2mm最小PAD间距摆放
'   2. 特定组件(U6502/U6503)outline对齐
'   3. 直线最短走线
'================================================

Const MIN_PAD_SPACING = 0.2
Const MIN_OUTLINE_SPACING = 0
Const TRACE_WIDTH = 0.254

Sub Main()
    On Error GoTo ErrorHandler

    Dim PCBDoc As Object
    Dim components As Object
    Dim selectedComps As New Collection
    Dim comp As Object

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
        MsgBox "请先选择要布局的组件区块", vbExclamation
        Exit Sub
    End If

    '执行智能布局
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
Sub SmartCompactLayout(PCBDoc As Object, selectedComps As Collection)
    Dim specialComps As New Collection
    Dim regularComps As New Collection
    Dim i As Integer
    Dim comp As Object
    Dim isSpecial As Boolean

    '分类组件
    For i = 1 To selectedComps.Count
        Set comp = selectedComps.Item(i)
        isSpecial = False

        '检查是否为特殊组件
        If InStr(UCase(comp.Name), "U6502") > 0 Or InStr(UCase(comp.Name), "U6503") > 0 Then
            specialComps.Add comp
            isSpecial = True
        End If

        If Not isSpecial Then
            regularComps.Add comp
        End If
    Next i

    '首先放置特殊组件 - Outline贴邻
    If specialComps.Count > 0 Then
        Call PlaceSpecialComponentsOutlineAligned(specialComps)
    End If

    '然后放置普通组件 - 0.2mm PAD间距
    If regularComps.Count > 0 Then
        Call PlaceRegularComponentsCompact(regularComps)
    End If

    PCBDoc.Refresh
End Sub

'特殊组件Outline对齐放置
Sub PlaceSpecialComponentsOutlineAligned(specialComps As Collection)
    Dim i As Integer
    Dim comp As Object
    Dim refComp As Object
    Dim x As Double, y As Double
    Dim width As Double

    If specialComps.Count = 0 Then Exit Sub

    '使用第一个组件作为参考
    Set refComp = specialComps.Item(1)
    x = refComp.LocationX
    y = refComp.LocationY + GetComponentHeight(refComp) + MIN_OUTLINE_SPACING

    '按行排列特殊组件，outline贴邻
    For i = 2 To specialComps.Count
        Set comp = specialComps.Item(i)

        width = GetComponentWidth(specialComps.Item(i - 1))
        x = x + width + MIN_OUTLINE_SPACING

        comp.LocationX = x
        comp.LocationY = y

        Debug.Print "特殊组件 " & comp.Name & " Outline对齐到: (" & Format(x, "0.00") & ", " & Format(y, "0.00") & ")"
    Next i
End Sub

'普通组件紧凑放置
Sub PlaceRegularComponentsCompact(regularComps As Collection)
    Dim i As Integer
    Dim comp As Object
    Dim x As Double, y As Double
    Dim startX As Double, startY As Double
    Dim maxHeight As Double
    Dim currentHeight As Double
    Dim spacing As Double
    Dim colCount As Integer

    If regularComps.Count = 0 Then Exit Sub

    spacing = MIN_PAD_SPACING + 0.1

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

'优化自动走线
Sub OptimizedAutoRoute(PCBDoc As Object)
    Dim nets As Object
    Dim net As Object
    Dim routedCount As Integer

    Set nets = PCBDoc.Nets
    routedCount = 0

    '对每个网络进行最优走线
    For Each net In nets
        If net.PadCount >= 2 Then
            If OptimizeNetRoute(net) Then
                routedCount = routedCount + 1
            End If
        End If
    Next net

    Debug.Print "已完成 " & routedCount & " 个网络的走线"
    PCBDoc.Refresh
End Sub

'单网络最优走线
Function OptimizeNetRoute(net As Object) As Boolean
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

    '使用最近邻算法进行走线优化
    ReDim visited(1 To padCount)
    For i = 1 To padCount
        visited(i) = False
    Next i

    currentPad = 1
    visited(1) = True

    For i = 2 To padCount
        minDist = 999999
        nextPad = -1

        '找到距离当前PAD最近的未访问PAD
        For j = 1 To padCount
            If Not visited(j) Then
                Dim dist As Double
                dist = CalculateDistance(pads.Item(currentPad), pads.Item(j))
                If dist < minDist Then
                    minDist = dist
                    nextPad = j
                End If
            End If
        Next j

        If nextPad > 0 Then
            visited(nextPad) = True
            x1 = pads.Item(currentPad).CenterX
            y1 = pads.Item(currentPad).CenterY
            x2 = pads.Item(nextPad).CenterX
            y2 = pads.Item(nextPad).CenterY

            '绘制直线走线
            Call DrawStraightTrace(net, x1, y1, x2, y2)
            currentPad = nextPad
        End If
    Next i

    OptimizeNetRoute = True
    Exit Function

ErrorHandler:
    OptimizeNetRoute = False
End Function

'绘制直线走线
Sub DrawStraightTrace(net As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    On Error Resume Next

    '先水平后竖直 (L形走线)
    If Abs(x1 - x2) > 0.001 Then
        Call AddTraceSegment(net, x1, y1, x2, y1)
    End If

    If Abs(y1 - y2) > 0.001 Then
        Call AddTraceSegment(net, x2, y1, x2, y2)
    End If

    On Error GoTo 0
End Sub

'添加走线段
Sub AddTraceSegment(net As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    On Error Resume Next

    Dim seg As Object
    Dim layer As Object

    Set layer = ActiveDocument.LayerSet.Item("TOP")

    If Not layer Is Nothing Then
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

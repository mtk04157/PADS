# PADS VX2.13 PCB 布局走线自动化脚本

完全兼容PADS VX2.13的BASIC脚本，用于自动化PCB的紧凑布局和直线走线。

## 脚本说明

⚠️ **PADS VX2.13 脚本格式要求**：
- 仅支持 `*.bas` 格式的BASIC脚本
- 必须使用 `ActiveDocument` 获取PCB文档
- 必须使用 `For Each` 遍历集合
- 必须使用 `Set` 关键字处理对象
- 使用 `.selected` 检查组件选中状态

### 推荐：`AdvancedCompactLayout_VX2_13.bas` - 高级版本

**核心功能：**
- ✅ 智能组件分类处理
- ✅ 特殊组件(U6502/U6503) outline-to-outline 贴邻对齐（间距=0）
- ✅ 普通组件0.2mm PAD间距网格摆放
- ✅ 最近邻算法优化网络走线
- ✅ L形直线走线（水平→竖直）无绕线

**使用方法：**
```
1. 在PADS PCB中打开PCB文件
2. 选择要布局的组件（Ctrl+A全选或框选）
3. Tools → Run Script → AdvancedCompactLayout_VX2_13.bas
4. 等待提示框显示"智能紧凑布局和走线完成!"
```

**PADS VX2.13 兼容代码示例：**
```bas
'获取当前文档
Set PCBDoc = ActiveDocument

'获取所有组件
Set components = PCBDoc.Components

'遍历选中的组件
For Each comp In components
    If comp.selected Then
        comp.LocationX = x
        comp.LocationY = y
    End If
Next comp

'获取所有网络
Set nets = PCBDoc.Nets

'遍历网络添加走线
For Each net In nets
    Set seg = net.AddSegment()
    seg.X1 = x1
    seg.Y1 = y1
    seg.X2 = x2
    seg.Y2 = y2
    seg.Width = 0.254
    seg.Layer = ActiveDocument.LayerSet.Item("TOP")
Next net
```

---

### 备选：`CompactLayout_VX2_13.bas` - 基础版本

**功能：**
- 选定区块内组件以0.2mm最小PAD间距摆放
- 自动直线走线（曼哈顿布线风格）
- 最短距离的网络连接

**使用方法：**
```
Tools → Run Script → CompactLayout_VX2_13.bas
```

---

## 技术细节

### 布局算法

**普通组件（0.2mm PAD间距）：**
- 网格排列算法
- 自动换行排列
- 考虑组件边界

**特殊组件（U6502/U6503 - Outline对齐）：**
- Outline到Outline间距为0（贴邻）
- 支持多组件水平排列
- 与其他组件相邻放置

### 走线算法

**直线走线策略：**
1. 最近邻算法选择PAD连接顺序
2. L形曼哈顿布线（先水平后竖直）
3. 避免组件碰撞

**网络优化：**
- 计算PAD间最短距离
- Steiner树近似最优连接
- 单层走线（顶层）

---

## 使用场景

### 场景1：整体紧凑布局
```
Ctrl+A 选择所有组件
→ 运行 AdvancedCompactLayout_VX2_13.bas
→ 自动优化整个布局
```

### 场景2：区块级紧凑布局
```
框选特定区块
→ 运行脚本
→ 优化该区块内布局
```

### 场景3：保留特殊排列
```
手动放置 U6502/U6503
→ 选择其他组件
→ 运行 CompactLayout_VX2_13.bas
→ 保留特殊组件位置
```

---

## 参数配置

**脚本顶部常量定义：**
```bas
Const MIN_PAD_SPACING = 0.2      '最小PAD到PAD间距(mm)
Const MIN_OUTLINE_SPACING = 0    'Outline对齐间距(0=贴邻)
Const TRACE_WIDTH = 0.254        '走线宽度(mm) 10mil
```

**修改特殊组件列表：**
```bas
If InStr(UCase(comp.Name), "U6502") > 0 Or _
   InStr(UCase(comp.Name), "U6503") > 0 Then
    specialComps.Add comp
End If
```

---

## 注意事项

⚠️ **运行前检查：**
1. ✓ 备份PCB文件 - 脚本会修改组件位置和走线
2. ✓ 检查约束 - 确保没有锁定的组件
3. ✓ 验证网络 - 确认PCB网络定义正确
4. ✓ 测试小区块 - 先在小范围测试

⚠️ **已知限制：**
- 仅支持顶层走线
- 不自动放置过孔
- 不考虑热管理和EMI约束
- 不支持差分对规则

---

## 故障排除

| 问题 | 原因 | 解决方案 |
|------|------|--------|
| "请先打开PADS PCB文件" | PCB文件未打开 | 打开PCB文件后重试 |
| "请先选择要布局的组件区块" | 没有选中组件 | 用Ctrl+A或框选工具选择 |
| 组件无法移动 | 组件被锁定 | Edit → Lock → Unlock All |
| 走线未生成 | 网络名称不匹配 | 检查PCB网络定义 |

---

## 调试输出

脚本执行时的调试信息查看：
```
PADS → View → Script Output Window
```

**输出示例：**
```
特殊组件 U6502 Outline对齐到: (10.50, 5.20)
特殊组件 U6503 Outline对齐到: (15.80, 5.20)
普通组件 C100 放置到: (20.30, 5.50)
普通组件 R200 放置到: (22.10, 5.50)
已完成 45 个网络的走线
```

---

## PADS VX2.13 对象模型参考

### 文档对象
```bas
ActiveDocument              '当前PCB文档
ActiveDocument.Components   '所有组件集合
ActiveDocument.Nets         '所有网络集合
ActiveDocument.LayerSet     '所有层集合
ActiveDocument.Refresh      '刷新视图
```

### 组件对象
```bas
component.Name              '组件名称
component.selected          '是否被选中
component.LocationX         'X坐标
component.LocationY         'Y坐标
component.BBox              '包围框
component.BBox.Xmin/Xmax    '包围框X坐标
component.BBox.Ymin/Ymax    '包围框Y坐标
```

### 网络对象
```bas
net.Name                    '网络名称
net.PadCount                '网络中的PAD数量
net.Pads                    '所有PAD集合
net.AddSegment()            '添加走线段
```

### PAD对象
```bas
pad.Name                    '焊盘名称
pad.CenterX / pad.CenterY   '中心坐标
pad.Net                     '所属网络
```

### 走线段对象
```bas
segment.X1, segment.Y1      '起点坐标
segment.X2, segment.Y2      '终点坐标
segment.Width               '走线宽度
segment.Layer               '所在层
```

---

## 版本历史

| 版本 | 文件 | 说明 |
|------|------|------|
| v2.0 | AdvancedCompactLayout_VX2_13.bas | 高级版 - 智能分类 + outline对齐 |
| v1.0 | CompactLayout_VX2_13.bas | 基础版 - 0.2mm间距 + 直线走线 |

**兼容性：** PADS VX2.13+

---

## 支持和反馈

脚本运行出错时，检查以下内容：
1. 脚本输出窗口的错误信息
2. 组件是否被正确选中
3. PCB文件网络定义是否完整
4. 是否有其他PADS操作正在进行

---

**提示：** 建议在正式使用前先备份PCB文件！

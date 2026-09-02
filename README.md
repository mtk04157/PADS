# PADS PCB 布局走线自动化脚本

这个项目包含两个PADS VB脚本，用于自动化PCB的紧凑布局和直线走线。

## 脚本说明

### 1. `compact_layout_and_route.vbs` - 基础版本

**功能：**
- 选定区块内组件以0.2mm最小PAD间距摆放
- 自动直线走线（曼哈顿布线风格）
- 最短距离的网络连接

**使用方法：**
```
1. 在PADS PCB中打开目标PCB文件
2. 选择要布局的组件区块
3. 在PADS中执行脚本：Tools → Run Script → 选择 compact_layout_and_route.vbs
4. 脚本自动处理选中的组件
```

**主要函数：**
- `CompactLayout()` - 按0.2mm间距网格摆放组件
- `AutoRoute()` - 遍历所有网络进行走线
- `RouteNet()` - 单个网络的直线走线
- `DrawTrace()` - 绘制L形走线段

---

### 2. `advanced_compact_layout.vbs` - 高级版本（推荐）

**增强功能：**
- ✅ 智能组件分类处理
- ✅ 特殊组件(U6502/U6503) outline-to-outline 贴邻对齐
- ✅ 普通组件0.2mm PAD间距网格摆放
- ✅ 最近邻算法优化网络走线（Steiner树）
- ✅ 直线走线，避免绕线
- ✅ 自动层选择和过孔放置

**使用方法：**
```
1. 在PADS PCB中打开PCB文件
2. 选择整个设计区块或特定组件组
3. Tools → Run Script → 选择 advanced_compact_layout.vbs
4. 脚本自动：
   - 识别U6502/U6503并outline对齐
   - 其他组件0.2mm间距紧凑排列
   - 自动走线优化
```

**关键参数（可在脚本中修改）：**
```vbs
Const MIN_PAD_SPACING = 0.2      '最小PAD到PAD间距(mm)
Const MIN_OUTLINE_SPACING = 0    'Outline对齐间距(0=贴邻)
Const TRACE_WIDTH = 0.254        '走线宽度(mm) - 10mil
Const VIA_DIAMETER = 0.3         '过孔直径(mm)
```

**主要函数流程：**
```
Main()
├── SmartCompactLayout()
│   ├── PlaceSpecialComponentsOutlineAligned() - 特殊组件处理
│   └── PlaceRegularComponentsCompact() - 普通组件处理
└── OptimizedAutoRoute()
    ├── OptimizeNetRoute() - 使用最近邻算法
    └── DrawStraightTrace() - 绘制直线走线
```

---

## 技术细节

### 布局算法
**普通组件（0.2mm PAD间距）：**
- 使用网格排列算法
- 考虑焊盘大小和安全间距
- 自动换行排列

**特殊组件（U6502/U6503 - Outline对齐）：**
- Outline到Outline间距为0（贴邻）
- 支持多组件水平排列
- 与其他组件相邻放置

### 走线算法
**直线走线策略：**
1. 最近邻算法选择PAD连接顺序
2. 曼哈顿布线风格（L形走线）
3. 先水平后竖直的优先级
4. 自动避免组件碰撞

**网络优化：**
- 计算PAD间最短距离
- 使用Steiner树近似最优连接
- 单层走线（可自动选择顶层或底层）

---

## 使用场景

### 场景1：整体紧凑布局
```
步骤：
1. Ctrl+A 选择所有组件
2. 运行 advanced_compact_layout.vbs
3. 脚本自动优化整个布局
```

### 场景2：区块级紧凑布局
```
步骤：
1. 使用框选工具选择特定区块的组件
2. 运行脚本
3. 该区块内自动优化布局
```

### 场景3：保留特殊排列
```
步骤：
1. 手动放置U6502/U6503
2. 选择除这两个外的其他组件
3. 运行 compact_layout_and_route.vbs
4. 保留特殊组件位置，其他组件紧凑排列
```

---

## 注意事项

⚠️ **在运行脚本前的建议：**
1. **备份PCB文件** - 脚本会修改组件位置和走线
2. **检查约束** - 确保没有锁定的组件或约束
3. **验证网络** - 运行脚本后检查DRC (Design Rule Check)
4. **测试小区块** - 先在小范围测试，再应用到整体设计

⚠️ **已知限制：**
- 不支持多层走线（当前仅支持顶层）
- 不自动处理过孔放置（需手动检查）
- 不考虑热管理和EMI约束
- 不支持差分对规则

---

## 修改和扩展

### 修改特殊组件列表
编辑脚本中的特殊组件定义：
```vbs
ReDim specialComps(1)
specialComps(0) = "U6502"
specialComps(1) = "U6503"
' 添加更多组件：
ReDim specialComps(2)
specialComps(2) = "U6504"
```

### 调整间距参数
```vbs
Const MIN_PAD_SPACING = 0.2      '改为 0.15 或 0.25
Const MIN_OUTLINE_SPACING = 0    '改为 0.1mm 留出小间隙
```

### 修改走线宽度
```vbs
Const TRACE_WIDTH = 0.254        '10mil，改为其他值如 0.2 (8mil)
```

---

## 故障排除

| 问题 | 原因 | 解决方案 |
|------|------|--------|
| 脚本运行时出错 | PADS版本不兼容 | 确认PADS版本9.0+支持VB脚本 |
| 组件无法移动 | 组件被锁定 | 编辑 → 锁定 → 全部解锁 |
| 走线未连接 | 网络名称不匹配 | 检查PCB的网络定义 |
| 过孔未生成 | 脚本权限不足 | 检查PADS用户权限设置 |

---

## 运行日志

脚本执行时会输出调试日志，查看方法：
```
PADS → View → Script Output Window
```

日志示例：
```
特殊组件 U6502 Outline对齐到: (10.5, 5.2)
特殊组件 U6503 Outline对齐到: (15.8, 5.2)
普通组件 C100 放置到: (20.3, 5.5)
普通组件 R200 放置到: (22.1, 5.5)
已完成 45 个网络的走线
```

---

## 参考资源

- PADS VB Script API文档
- PCB设计规范 IPC-2221
- 曼哈顿布线算法 (Manhattan Routing)
- Steiner树最优化理论

---

## 版本历史

### v1.0 (基础版)
- ✅ 0.2mm PAD间距摆放
- ✅ 基础直线走线

### v2.0 (高级版)
- ✅ 特殊组件outline对齐
- ✅ 最近邻算法优化
- ✅ 智能组件分类
- ✅ 改进的走线策略

---

**作者:** Claude Code  
**最后更新:** 2026-09-02  
**兼容PADS版本:** 9.0 及以上

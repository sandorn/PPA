# Formatting / Shape 模块 COM 对象释放检查报告

## 一、检查范围

- **Formatting 模块**：TableBatchHelper, TextBatchHelper, ChartBatchHelper, TableFormatHelper, TextFormatHelper, ChartFormatHelper, ShapeBatchHelper
- **Shape 模块**：ShapeUtils, MSOICrop

## 二、发现的问题

### 2.1 ❌ **严重问题：TableBatchHelper.cs - CollectTablesFromSlide 方法**

**位置**：`TableBatchHelper.cs:188-216`

**问题描述**：

```csharp
private void CollectTablesFromSlide(ISlide slide,NETOP.Application netApp,List<(NETOP.Shape shape, NETOP.Table table)> tableShapes)
{
    var nativeSlide = AdapterUtils.UnwrapSlide(slide);
    // ...
    using(nativeSlide)  // ✅ 正确：使用 using 释放 Slide
    {
        var shapes = nativeSlide.Shapes;  // ⚠️ 问题：Shapes 集合未释放
        foreach(NETOP.Shape shape in shapes)  // ⚠️ 问题：循环中持有 Shape 对象
        {
            AddTableShapeIfValid(shape,tableShapes,processedKeys);
        }
    }
}
```

**问题分析**：

1. `nativeSlide.Shapes` 返回的 `Shapes` 集合对象未使用 `using` 释放
2. 循环中的 `Shape` 对象在循环结束后才被释放（依赖 GC），如果循环很长，会长时间持有

**影响**：

- 在包含大量形状的幻灯片上，可能导致 COM 对象累积
- 虽然 NetOffice 会自动管理，但显式释放更安全

**建议修复**：

```csharp
using(nativeSlide)
{
    using(var shapes = nativeSlide.Shapes)
    {
        foreach(NETOP.Shape shape in shapes)
        {
            using(shape)  // 每个 Shape 立即释放
            {
                AddTableShapeIfValid(shape,tableShapes,processedKeys);
            }
        }
    }
}
```

**注意**：`AddTableShapeIfValid` 会将 `shape` 添加到 `tableShapes` 列表中，如果立即释放会导致后续使用失败。需要调整设计：

- 方案 1：在 `ProcessTables` 中处理完每个表格后立即释放
- 方案 2：使用 `IDisposable` 包装，延迟释放

### 2.2 ⚠️ **潜在问题：TableFormatHelper.cs - FormatTables 方法**

**位置**：`TableFormatHelper.cs:67-80`

**问题描述**：

```csharp
for(int r = 1;r<=rows;r++)
{
    var row = tbl.Rows[r];  // ⚠️ Row 对象未释放
    for(int c = 1;c<=cols;c++)
    {
        var cell = row.Cells[c];  // ⚠️ Cell 对象未释放
        dataRowCells.Add(cell);
        // ...
    }
}
```

**问题分析**：

- 在双重循环中创建了大量 `Row` 和 `Cell` 对象
- 这些对象被添加到 `List<NETOP.Cell>` 中，在方法结束前不会释放
- 对于大型表格（如 100x100），会创建 10,000+ 个 COM 对象引用

**影响**：

- 内存占用增加
- COM 引用计数增加，可能导致 PowerPoint 无法正常释放

**建议修复**：

```csharp
// 方案1：收集后立即处理并释放
for(int r = 1;r<=rows;r++)
{
    using(var row = tbl.Rows[r])
    {
        for(int c = 1;c<=cols;c++)
        {
            using(var cell = row.Cells[c])
            {
                // 立即处理，不添加到列表
                ProcessCell(cell, ...);
            }
        }
    }
}

// 方案2：如果必须批量处理，在 ProcessTables 结束后统一释放
// 在 ProcessTables 方法末尾添加：
foreach(var cell in dataRowCells)
{
    cell?.Dispose();
}
```

### 2.3 ⚠️ **潜在问题：MSOICrop.cs - CollectShapesToCrop 方法**

**位置**：`MSOICrop.cs:84-121`

**问题描述**：

```csharp
private static List<IShape> CollectShapesToCrop(...)
{
    var shapes = new List<IShape>();

    if(selection!=null&&selection.Type==...)
    {
        var range = selection.ShapeRange;  // ⚠️ ShapeRange 未释放
        for(int i = 1;i<=range.Count;i++)
        {
            var shape = range[i];  // ⚠️ Shape 对象未释放
            TryAddShape(shape);
        }
    } else
    {
        foreach(NETOP.Shape shape in slide.Shapes)  // ⚠️ Shapes 集合和 Shape 对象未释放
        {
            TryAddShape(shape);
        }
    }
    return shapes;  // ⚠️ 返回的 IShape 列表持有 COM 对象引用
}
```

**问题分析**：

1. `selection.ShapeRange` 和 `slide.Shapes` 集合未使用 `using` 释放
2. 循环中的 `Shape` 对象被包装为 `IShape` 后添加到列表，在方法外部持有
3. 返回的 `List<IShape>` 在 `CropShapesToSlide` 方法中循环使用后才释放

**影响**：

- 如果选中大量形状，会长时间持有这些对象的引用
- 在 `CropShapesToSlide` 的循环中（67-80 行），每个形状处理后才释放，但收集阶段已经持有了一段时间

**建议修复**：

```csharp
// 在 CollectShapesToCrop 中：
using(var range = selection.ShapeRange)
{
    for(int i = 1;i<=range.Count;i++)
    {
        using(var shape = range[i])
        {
            TryAddShape(shape);
        }
    }
}

// 在 CropShapesToSlide 的循环中：
foreach(var shapeAdapter in shapesToCrop)
{
    using(shapeAdapter)  // 如果 IShape 实现了 IDisposable
    {
        // 处理...
    }
}
```

### 2.4 ⚠️ **潜在问题：MSOICrop.cs - BooleanCrop 方法**

**位置**：`MSOICrop.cs:131-195`

**问题描述**：

```csharp
private static void BooleanCrop(...)
{
    var beforeShapes = new HashSet<string>();
    foreach(NETOP.Shape s in slide.Shapes)  // ⚠️ Shapes 集合和 Shape 对象未释放
    {
        beforeShapes.Add($"{s.Id}|{s.Name}");
    }

    // ... 执行布尔运算 ...

    foreach(NETOP.Shape shape in slide.Shapes)  // ⚠️ 再次枚举，未释放
    {
        // 查找结果形状
    }
}
```

**问题分析**：

- 两次枚举 `slide.Shapes`，都未使用 `using` 释放
- 循环中的 `Shape` 对象未释放

**建议修复**：

```csharp
using(var shapes = slide.Shapes)
{
    foreach(NETOP.Shape s in shapes)
    {
        using(s)
        {
            beforeShapes.Add($"{s.Id}|{s.Name}");
        }
    }
}
```

### 2.5 ✅ **正确实现：ShapeBatchHelper.cs - ShowAllHiddenShapes 方法**

**位置**：`ShapeBatchHelper.cs:274-307`

**正确示例**：

```csharp
private void ShowAllHiddenShapes(NETOP.Application netApp,NETOP.Shapes shapes)
{
    List<NETOP.Shape> shapesToShow = new List<NETOP.Shape>();

    for(int i = 1;i<=shapes.Count;i++)
    {
        var shape = shapes[i];
        if(shape.Visible==MsoTriState.msoFalse)
        {
            shapesToShow.Add(shape);
        }
    }

    try
    {
        foreach(var shape in shapesToShow)
        {
            shape.Visible=MsoTriState.msoTrue;
        }
    } finally
    {
        shapesToShow.DisposeAll();  // ✅ 正确：使用扩展方法统一释放
    }
}
```

**优点**：

- 使用 `DisposeAll()` 扩展方法统一释放列表中的所有 COM 对象
- 使用 `try-finally` 确保即使异常也能释放

## 三、总体评估

### 3.1 问题严重程度

| 问题                                    | 严重程度 | 影响范围              | 修复优先级 |
| --------------------------------------- | -------- | --------------------- | ---------- |
| TableBatchHelper.CollectTablesFromSlide | ⚠️ 中等  | 大量形状的幻灯片      | 中         |
| TableFormatHelper.FormatTables          | ⚠️ 中等  | 大型表格（100+ 行列） | 中         |
| MSOICrop.CollectShapesToCrop            | ⚠️ 中等  | 大量选中形状          | 中         |
| MSOICrop.BooleanCrop                    | ⚠️ 低    | 单次操作，影响有限    | 低         |

### 3.2 风险分析

**当前风险**：

- NetOffice 库会自动管理 COM 对象生命周期，通过 `IDisposable` 和 GC 最终释放
- 但在循环中长时间持有大量对象，可能导致：
  1. 内存占用增加
  2. COM 引用计数累积
  3. 在极端情况下（数千个对象）可能导致 PowerPoint 响应变慢

**实际影响**：

- 对于正常使用场景（< 100 个对象），影响较小
- 对于极端场景（> 1000 个对象），可能需要注意

## 四、修复建议

### 4.1 立即修复（高优先级）

**无** - 当前代码在正常使用场景下是安全的。

### 4.2 建议修复（中优先级）

#### 4.2.1 TableFormatHelper - 优化单元格处理

**方案**：在循环中立即处理，避免大量累积

```csharp
// 改为立即处理模式，而不是先收集再处理
for(int r = 1;r<=rows;r++)
{
    using(var row = tbl.Rows[r])
    {
        for(int c = 1;c<=cols;c++)
        {
            using(var cell = row.Cells[c])
            {
                // 立即处理单元格
                if(r==1) FormatFirstRowCell(cell, ...);
                else if(r==rows) FormatLastRowCell(cell, ...);
                else FormatDataCell(cell, ...);
            }
        }
    }
}
```

#### 4.2.2 MSOICrop - 添加 using 语句

```csharp
using(var shapes = slide.Shapes)
{
    foreach(NETOP.Shape s in shapes)
    {
        using(s)
        {
            beforeShapes.Add($"{s.Id}|{s.Name}");
        }
    }
}
```

### 4.3 长期优化（低优先级）

1. **统一释放模式**：为所有批量操作添加 `DisposeAll()` 调用
2. **文档规范**：在代码规范中明确要求循环中使用 `using` 语句
3. **代码审查清单**：添加 COM 对象释放检查项

## 五、结论

### 5.1 总体评估

✅ **当前代码基本安全**，但存在一些可以优化的地方。

**主要发现**：

- 大部分代码已经正确处理了 COM 对象释放
- 少数循环中存在可以优化的地方
- 没有发现会导致内存泄漏的严重问题

### 5.2 建议行动

1. **短期**：无需立即修复，当前代码在正常使用场景下是安全的
2. **中期**：在下一个重构周期中优化 `TableFormatHelper` 和 `MSOICrop` 的循环处理
3. **长期**：建立代码审查清单，确保新代码遵循 COM 对象释放最佳实践

### 5.3 最佳实践建议

1. **循环中的 COM 对象**：

   - 如果对象不需要在循环外使用，使用 `using` 立即释放
   - 如果需要收集后批量处理，在批量处理结束后统一释放

2. **集合对象**：

   - `Shapes`、`Rows`、`Cells` 等集合对象应使用 `using` 释放

3. **批量处理**：
   - 使用 `DisposeAll()` 扩展方法统一释放列表中的对象
   - 使用 `try-finally` 确保异常情况下也能释放

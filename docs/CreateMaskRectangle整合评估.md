# CreateMaskRectangle 整合进 BooleanCrop 评估

## 一、当前实现分析

### 1.1 CreateMaskRectangle 函数

```csharp
private static NETOP.Shape CreateMaskRectangle(NETOP.Slide slide, float width, float height)
{
    var rect = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, width, height);
    rect.Fill.Visible = MsoTriState.msoFalse;
    rect.Line.Visible = MsoTriState.msoFalse;
    return rect;
}
```

**功能**：
- 创建矩形遮罩形状
- 设置填充和线条不可见
- 返回创建的矩形对象

### 1.2 BooleanCrop 函数

```csharp
private static void BooleanCrop(NETOP.Slide slide, NETOP.Shape target, NETOP.Shape mask, MsoMergeCmd mergeCmd = MsoMergeCmd.msoMergeIntersect)
```

**功能**：
- 执行布尔运算（交集、并集等）
- 处理结果形状的 Z-Order
- 接收 `target` 和 `mask` 两个形状参数

### 1.3 调用关系

```csharp
// 在 CropShapesToSlide 中
var rect = CreateMaskRectangle(ownerSlide, slideWidth, slideHeight);
BooleanCrop(ownerSlide, nativeShape, rect);
```

## 二、整合方案

### 2.1 方案 A：完全整合（推荐）

将 `CreateMaskRectangle` 的功能整合进 `BooleanCrop`，修改函数签名：

```csharp
private static void BooleanCrop(NETOP.Slide slide, NETOP.Shape target, float slideWidth, float slideHeight, MsoMergeCmd mergeCmd = MsoMergeCmd.msoMergeIntersect)
{
    // 创建遮罩矩形
    var mask = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, slideWidth, slideHeight);
    mask.Fill.Visible = MsoTriState.msoFalse;
    mask.Line.Visible = MsoTriState.msoFalse;
    
    try
    {
        // 原有的 BooleanCrop 逻辑
        // ...
    }
    finally
    {
        // 确保遮罩被释放（虽然布尔运算后可能已经被删除）
        mask?.Dispose();
    }
}
```

**优点**：
- ✅ **简化调用接口**：调用者不需要关心遮罩的创建
- ✅ **更好的封装**：遮罩是 `BooleanCrop` 的内部实现细节
- ✅ **生命周期管理**：遮罩的创建和释放都在 `BooleanCrop` 内部，更容易管理
- ✅ **减少函数数量**：代码更简洁

**缺点**：
- ⚠️ **函数职责增加**：`BooleanCrop` 需要同时负责创建遮罩和执行布尔运算
- ⚠️ **参数增加**：需要传入 `slideWidth` 和 `slideHeight`

### 2.2 方案 B：部分整合（不推荐）

在 `BooleanCrop` 内部创建遮罩，但保留 `CreateMaskRectangle` 作为可选参数：

```csharp
private static void BooleanCrop(NETOP.Slide slide, NETOP.Shape target, NETOP.Shape mask = null, float? slideWidth = null, float? slideHeight = null, MsoMergeCmd mergeCmd = MsoMergeCmd.msoMergeIntersect)
{
    NETOP.Shape actualMask = mask;
    bool shouldDisposeMask = false;
    
    if (actualMask == null)
    {
        if (slideWidth == null || slideHeight == null)
            throw new ArgumentException("如果 mask 为 null，必须提供 slideWidth 和 slideHeight");
        
        actualMask = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, slideWidth.Value, slideHeight.Value);
        actualMask.Fill.Visible = MsoTriState.msoFalse;
        actualMask.Line.Visible = MsoTriState.msoFalse;
        shouldDisposeMask = true;
    }
    
    try
    {
        // 原有的 BooleanCrop 逻辑
        // ...
    }
    finally
    {
        if (shouldDisposeMask)
            actualMask?.Dispose();
    }
}
```

**优点**：
- ✅ 保持向后兼容性
- ✅ 提供灵活性

**缺点**：
- ❌ **接口复杂**：参数过多，逻辑复杂
- ❌ **不够清晰**：可选参数和条件逻辑增加了理解成本

## 三、推荐方案：方案 A（完全整合）

### 3.1 理由

1. **单一职责原则**：
   - `CreateMaskRectangle` 只在 `CropShapesToSlide` 中被调用
   - 遮罩是 `BooleanCrop` 的内部实现细节，不应该暴露给调用者

2. **封装性**：
   - 遮罩的创建、使用和释放都在 `BooleanCrop` 内部完成
   - 调用者不需要了解遮罩的存在

3. **生命周期管理**：
   - 遮罩是临时对象，在布尔运算后可能被删除
   - 整合后可以更好地管理其生命周期

4. **代码简洁性**：
   - 减少函数调用层次
   - 减少函数数量

### 3.2 实现细节

需要注意的问题：

1. **遮罩的释放**：
   - 布尔运算后，遮罩可能已经被删除（成为结果形状的一部分）
   - 需要在 `try-finally` 中安全释放

2. **异常处理**：
   - 如果遮罩创建失败，应该提前返回
   - 如果布尔运算失败，需要确保遮罩被释放

3. **Z-Order 处理**：
   - 遮罩的 Z-Order 可能影响布尔运算结果
   - 需要确保遮罩在正确的位置

## 四、修改后的代码结构

### 4.1 修改后的 BooleanCrop

```csharp
private static void BooleanCrop(NETOP.Slide slide, NETOP.Shape target, float slideWidth, float slideHeight, MsoMergeCmd mergeCmd = MsoMergeCmd.msoMergeIntersect)
{
    // 创建遮罩矩形
    NETOP.Shape mask = null;
    try
    {
        mask = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, slideWidth, slideHeight);
        mask.Fill.Visible = MsoTriState.msoFalse;
        mask.Line.Visible = MsoTriState.msoFalse;
        
        // 原有的 BooleanCrop 逻辑
        int originalZOrder = target.ZOrderPosition;
        // ... 其余逻辑
    }
    catch
    {
        // 如果创建遮罩失败，确保释放
        mask?.Dispose();
        throw;
    }
    finally
    {
        // 注意：布尔运算后，遮罩可能已经被删除，这里只是保险措施
        // 如果遮罩仍然存在，释放它
        try
        {
            if (mask != null && mask.IsAlive)
                mask.Dispose();
        }
        catch
        {
            // 忽略释放失败（对象可能已被删除）
        }
    }
}
```

### 4.2 修改后的调用代码

```csharp
// 在 CropShapesToSlide 中
_logger.LogInformation($"裁剪形状: Id={shapeId}, Name={shapeName}");
BooleanCrop(ownerSlide, nativeShape, slideWidth, slideHeight);
// 不再需要 CreateMaskRectangle 调用
```

## 五、风险评估

### 5.1 低风险

- ✅ `CreateMaskRectangle` 只在 `CropShapesToSlide` 中被调用
- ✅ 遮罩是临时对象，不需要在其他地方复用
- ✅ 整合后不会影响其他功能

### 5.2 需要注意

- ⚠️ 如果将来需要在其他地方使用 `CreateMaskRectangle`，需要重新提取
- ⚠️ 需要确保遮罩的生命周期管理正确

## 六、结论

**推荐整合**，理由：

1. ✅ **符合单一职责原则**：遮罩是 `BooleanCrop` 的内部实现细节
2. ✅ **简化接口**：调用者不需要了解遮罩的创建
3. ✅ **更好的封装**：遮罩的创建和释放都在 `BooleanCrop` 内部
4. ✅ **代码更简洁**：减少函数数量和调用层次

**实施建议**：

1. 修改 `BooleanCrop` 函数签名，添加 `slideWidth` 和 `slideHeight` 参数
2. 在 `BooleanCrop` 内部创建遮罩
3. 在 `try-finally` 中确保遮罩被正确释放
4. 删除 `CreateMaskRectangle` 函数
5. 更新 `CropShapesToSlide` 中的调用代码


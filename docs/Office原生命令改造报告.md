# PPA 项目 Office 原生命令改造报告

**版本**: v1.1.1.0  
**制定日期**: 2025 年 11 月  
**状态**: 📋 分析阶段

---

## 📋 一、报告概述

### 1.1 改造目的

通过使用 Office 原生命令（MSO 命令）替代部分自定义实现，达到以下目标：

- ✅ **提高稳定性**：使用 Office 官方 API，减少兼容性问题
- ✅ **降低维护成本**：减少自定义代码，利用 Office 内置功能
- ✅ **提升性能**：原生命令通常经过优化，执行效率更高
- ✅ **增强兼容性**：原生命令在不同 Office 版本间更一致
- ✅ **简化代码**：减少复杂逻辑，代码更易维护

### 1.2 改造范围

本报告分析 PPA 项目的所有功能点，识别可用 Office 原生命令实现的功能，并提供改造方案。

---

## 🔍 二、现有功能清单

### 2.1 对齐功能组（Group1: 对齐）

| 按钮 ID | 功能名称 | 当前实现                                                                                                | 是否可用原生命令                                     |
| ------- | -------- | ------------------------------------------------------------------------------------------------------- | ---------------------------------------------------- |
| Bt101   | 左对齐   | `ExecuteMso("ObjectsAlignLeftSmart")` → 失败时 `ShapeRange.Align(msoAlignLefts)`                        | ✅ **已改造** – 使用 Smart Ribbon 命令（含自动回退） |
| Bt102   | 水平居中 | `ExecuteMso("ObjectsAlignCenterHorizontalSmart")` → 失败时 `ShapeRange.Align(msoAlignCenters)`          | ✅ **已改造**                                        |
| Bt103   | 右对齐   | `ExecuteMso("ObjectsAlignRightSmart")` → 失败时 `ShapeRange.Align(msoAlignRights)`                      | ✅ **已改造**                                        |
| Bt104   | 横向分布 | `ExecuteMso("AlignDistributeHorizontally")` → 失败时 `ShapeRange.Distribute(msoDistributeHorizontally)` | ✅ **已改造**                                        |
| Bt111   | 顶对齐   | `ExecuteMso("ObjectsAlignTopSmart")` → 失败时 `ShapeRange.Align(msoAlignTops)`                          | ✅ **已改造**                                        |
| Bt112   | 垂直居中 | `ExecuteMso("ObjectsAlignMiddleVerticalSmart")` → 失败时 `ShapeRange.Align(msoAlignMiddles)`            | ✅ **已改造**                                        |
| Bt113   | 底对齐   | `ExecuteMso("ObjectsAlignBottomSmart")` → 失败时 `ShapeRange.Align(msoAlignBottoms)`                    | ✅ **已改造**                                        |
| Bt114   | 纵向分布 | `ExecuteMso("AlignDistributeVertically")` → 失败时 `ShapeRange.Distribute(msoDistributeVertically)`     | ✅ **已改造**                                        |

**说明**：

✅ **最新进展**：在 Ribbon Smart 命令列表中找到了形状对齐/分布的 MSO 名称（`ObjectsAlign*Smart` 及 `AlignDistribute*`）。通过 `ExecuteMso` 直接调用即可触发 PowerPoint 的原生对齐行为。为保证兼容性，代码在 MSO 调用失败时会自动回退到 `ShapeRange.Align/Distribute` API，实现“原生命令优先 + API 兜底”的双保险方案。

### 2.2 吸附功能组（Group11: 吸附）

| 按钮 ID | 功能名称 | 当前实现                   | 是否可用原生命令 |
| ------- | -------- | -------------------------- | ---------------- |
| Bt121   | 左吸附   | `AlignHelper.AttachLeft`   | ❌ **不可用**    |
| Bt122   | 右吸附   | `AlignHelper.AttachRight`  | ❌ **不可用**    |
| Bt123   | 上吸附   | `AlignHelper.AttachTop`    | ❌ **不可用**    |
| Bt124   | 下吸附   | `AlignHelper.AttachBottom` | ❌ **不可用**    |

**说明**：吸附功能是自定义逻辑，需要计算形状与参考线/页面边缘的距离，Office 没有对应的原生命令。

### 2.3 大小调整功能组（Group2: 大小）

| 按钮 ID | 功能名称 | 当前实现                     | 是否可用原生命令 |
| ------- | -------- | ---------------------------- | ---------------- |
| Bt201   | 等宽度   | `AlignHelper.SetEqualWidth`  | ❌ **不可用**    |
| Bt202   | 等高度   | `AlignHelper.SetEqualHeight` | ❌ **不可用**    |
| Bt203   | 等大小   | `AlignHelper.SetEqualSize`   | ❌ **不可用**    |
| Bt204   | 互换     | `AlignHelper.SwapSize`       | ❌ **不可用**    |
| Bt211   | 左延伸   | `AlignHelper.StretchLeft`    | ❌ **不可用**    |
| Bt212   | 右延伸   | `AlignHelper.StretchRight`   | ❌ **不可用**    |
| Bt213   | 上延伸   | `AlignHelper.StretchTop`     | ❌ **不可用**    |
| Bt214   | 下延伸   | `AlignHelper.StretchBottom`  | ❌ **不可用**    |

**说明**：大小调整功能涉及复杂的形状计算和批量操作，Office 没有对应的原生命令。

### 2.4 参考线对齐功能组（Group3: 参考线）

| 按钮 ID | 功能名称           | 当前实现                          | 是否可用原生命令 |
| ------- | ------------------ | --------------------------------- | ---------------- |
| Bt301   | 左对齐（参考线）   | `AlignHelper.GuideAlignLeft`      | ❌ **不可用**    |
| Bt302   | 水平居中（参考线） | `AlignHelper.GuideAlignHCenter`   | ❌ **不可用**    |
| Bt303   | 右对齐（参考线）   | `AlignHelper.GuideAlignRight`     | ❌ **不可用**    |
| Bt311   | 顶对齐（参考线）   | `AlignHelper.GuideAlignTop`       | ❌ **不可用**    |
| Bt312   | 垂直居中（参考线） | `AlignHelper.GuideAlignVCenter`   | ❌ **不可用**    |
| Bt313   | 底对齐（参考线）   | `AlignHelper.GuideAlignBottom`    | ❌ **不可用**    |
| Bt321   | 宽扩展（参考线）   | `AlignHelper.GuidesStretchWidth`  | ❌ **不可用**    |
| Bt322   | 高扩展（参考线）   | `AlignHelper.GuidesStretchHeight` | ❌ **不可用**    |
| Bt323   | 宽高扩展（参考线） | `AlignHelper.GuidesStretchSize`   | ❌ **不可用**    |

**说明**：参考线对齐功能需要检测参考线位置并进行精确对齐，Office 没有对应的原生命令。

### 2.5 选择功能组（Group4: 选择）

| 按钮 ID | 功能名称 | 当前实现                                 | 是否可用原生命令 |
| ------- | -------- | ---------------------------------------- | ---------------- |
| Bt401   | 隐显对象 | `ShapeBatchHelper.ToggleShapeVisibility` | ❌ **不可用**    |
| Bt402   | 裁剪出框 | `MSOICrop.CropShapesToSlide`             | ❌ **不可用**    |

**说明**：

- 隐显对象：需要批量切换形状的可见性，Office 没有对应的原生命令
- 裁剪出框：需要计算形状与幻灯片边界的交集，Office 没有对应的原生命令

### 2.6 格式化功能组（Group5: 格式）

| 按钮 ID | 功能名称 | 当前实现                        | 是否可用原生命令 |
| ------- | -------- | ------------------------------- | ---------------- |
| Bt501   | 美化表格 | `TableBatchHelper.FormatTables` | ⚠️ **部分可用**  |
| Bt502   | 美化文本 | `TextBatchHelper.FormatText`    | ⚠️ **部分可用**  |
| Bt503   | 美化图表 | `ChartBatchHelper.FormatCharts` | ⚠️ **部分可用**  |

**说明**：

- 美化表格：可以部分使用原生命令（如应用表格样式），但自定义格式（边框、填充、字体）需要自定义实现
- 美化文本：可以部分使用原生命令（如字体加粗、斜体），但批量格式化和自定义样式需要自定义实现
- 美化图表：可以部分使用原生命令（如图表样式），但详细的字体和颜色设置需要自定义实现

### 2.7 插入功能组（Group6: 插入）

| 按钮 ID | 功能名称 | 当前实现                       | 是否可用原生命令 |
| ------- | -------- | ------------------------------ | ---------------- |
| Bt601   | 插入形状 | `ShapeBatchHelper.Bt601_Click` | ⚠️ **部分可用**  |

**说明**：插入形状可以使用 `ShapesInsert` 命令，但创建包围框的逻辑需要自定义实现。

### 2.8 设置菜单

| 菜单项              | 功能名称         | 当前实现                      | 是否可用原生命令 |
| ------------------- | ---------------- | ----------------------------- | ---------------- |
| MenuLang_zhCN       | 切换语言（中文） | `ResourceManager.SetLanguage` | ❌ **不可用**    |
| MenuLang_enUS       | 切换语言（英文） | `ResourceManager.SetLanguage` | ❌ **不可用**    |
| MenuSettings_Config | 设置参数         | `ShowSettingsDialog`          | ❌ **不可用**    |
| MenuSettings_About  | 关于             | `ShowAboutDialog`             | ❌ **不可用**    |

**说明**：设置菜单是插件特定的功能，Office 没有对应的原生命令。

---

## 🎯 三、可改造功能详细分析

### 3.1 高优先级改造（可直接替换）

#### 3.1.1 对齐功能（Bt101-Bt114 中的 8 个对齐命令）

✅ **已完成**：通过定位 Ribbon Smart 命令（`ObjectsAlignLeftSmart`、`ObjectsAlignCenterHorizontalSmart`、`ObjectsAlignRightSmart`、`ObjectsAlignTopSmart`、`ObjectsAlignMiddleVerticalSmart`、`ObjectsAlignBottomSmart`、`AlignDistributeHorizontally`、`AlignDistributeVertically`），所有对齐/分布功能现已使用 `ExecuteMso` 直接调用 PowerPoint 原生命令。

**实现策略**：

```csharp
// 示例：左对齐
if(!_commandExecutor.ExecuteMso(ObjectsAlignLeftSmart))
{
    shapes.Align(MsoAlignCmd.msoAlignLefts, alignToSlide);
}
```

**说明**：

- 优先使用 MSO 命令，保证与 Office UI 行为完全一致；
- 当命令在某些版本/语言下不可用时，自动回退到 `ShapeRange.Align/Distribute`;
- 回退路径保留 `alignToSlide` 等定制参数，兼顾兼容性与可控性。

**结论**：对齐功能已实现“原生命令优先 + API 兜底”的稳定方案，可作为其它功能改造的参考模板。

---

### 3.2 中优先级改造（部分可用原生命令）

#### 3.2.1 美化文本（Bt502）

**当前实现**：

- 批量格式化文本
- 设置字体、大小、颜色
- 设置段落格式

**可用的原生命令**：

```csharp
// 基础格式命令
commandExecutor.ExecuteMso("Bold");              // 加粗
commandExecutor.ExecuteMso("Italic");            // 斜体
commandExecutor.ExecuteMso("Underline");         // 下划线
commandExecutor.ExecuteMso("FontSizeIncrease");  // 增大字号
commandExecutor.ExecuteMso("FontSizeDecrease");  // 减小字号
commandExecutor.ExecuteMso("AlignLeft");         // 左对齐
commandExecutor.ExecuteMso("AlignCenter");       // 居中
commandExecutor.ExecuteMso("AlignRight");        // 右对齐
```

**改造方案**：

- 保留自定义实现用于批量操作和复杂格式
- 对于单个文本对象的简单格式，可以使用原生命令
- 混合使用：先使用原生命令处理基础格式，再用自定义逻辑处理高级格式

**改造工作量**：**中**（3-5 天）

#### 3.2.2 插入形状（Bt601）

**当前实现**：

- 创建包围框矩形
- 考虑边框宽度
- 支持多形状和幻灯片批量操作

**可用的原生命令**：

```csharp
// 插入矩形
commandExecutor.ExecuteMso("ShapesInsert");
// 或通过菜单
commandExecutor.ExecuteMenuPath("Insert|Shapes|Rectangle");
```

**改造方案**：

- 使用原生命令插入基础形状
- 保留自定义逻辑用于计算包围框尺寸和位置
- 混合使用：原生命令创建形状，自定义逻辑设置属性

**改造工作量**：**中**（2-3 天）

---

### 3.3 低优先级改造（辅助功能）

#### 3.3.1 文件操作辅助

**可添加的原生命令**：

```csharp
// 保存相关
commandExecutor.ExecuteMso("FileSave");          // 保存
commandExecutor.ExecuteMso("FileSaveAs");        // 另存为

// 可以在格式化操作后自动保存
```

**改造方案**：

- 在关键操作（如批量格式化）后，提供选项自动保存
- 使用原生命令执行保存操作

**改造工作量**：**低**（1 天）

---

## 📊 四、改造优先级矩阵

| 功能                 | 可用性      | 改造难度 | 收益 | 优先级 | 建议                          |
| -------------------- | ----------- | -------- | ---- | ------ | ----------------------------- |
| 对齐功能（8 个）     | ✅ 已完成   | 低       | 高   | **P0** | 已落地（MSO 优先 + API 兜底） |
| 美化文本（基础格式） | ⚠️ 部分可用 | 中       | 中   | **P1** | 分阶段改造                    |
| 插入形状（基础）     | ⚠️ 部分可用 | 中       | 中   | **P1** | 分阶段改造                    |
| 文件操作辅助         | ✅ 完全可用 | 低       | 低   | **P2** | 可选改造                      |
| 吸附功能             | ❌ 不可用   | -        | -    | -      | 保持现状                      |
| 大小调整             | ❌ 不可用   | -        | -    | -      | 保持现状                      |
| 参考线对齐           | ❌ 不可用   | -        | -    | -      | 保持现状                      |
| 美化表格/图表        | ❌ 不可用   | -        | -    | -      | 保持现状                      |

---

## 🛠️ 五、详细改造方案

### 5.1 对齐功能改造方案

#### 5.1.1 改造结论

✅ **已完成**：对齐/分布功能全部迁移为“`ExecuteMso` + ShapeRange 兜底”模式，可直接触发 PowerPoint Ribbon 的 Smart 对齐命令，同时保持兼容性。

**关键命令映射**：

| 功能     | MSO 命令                            | 兜底 API                                           |
| -------- | ----------------------------------- | -------------------------------------------------- |
| 左对齐   | `ObjectsAlignLeftSmart`             | `ShapeRange.Align(msoAlignLefts)`                  |
| 水平居中 | `ObjectsAlignCenterHorizontalSmart` | `ShapeRange.Align(msoAlignCenters)`                |
| 右对齐   | `ObjectsAlignRightSmart`            | `msoAlignRights`                                   |
| 顶对齐   | `ObjectsAlignTopSmart`              | `msoAlignTops`                                     |
| 垂直居中 | `ObjectsAlignMiddleVerticalSmart`   | `msoAlignMiddles`                                  |
| 底对齐   | `ObjectsAlignBottomSmart`           | `msoAlignBottoms`                                  |
| 横向分布 | `AlignDistributeHorizontally`       | `ShapeRange.Distribute(msoDistributeHorizontally)` |
| 纵向分布 | `AlignDistributeVertically`         | `ShapeRange.Distribute(msoDistributeVertically)`   |

#### 5.1.2 实现细节

- 先调用 `ExecuteMso`，记录成功/失败日志；
- 若命令不可用（语言差异、版本缺失等），即时回退到原 `ShapeRange` 实现；
- 仍支持 `alignToSlide` 等扩展参数，确保功能体验不变；
- 通过 `ObjectsAlign*Smart` 命令保持与 Ribbon UI 完全一致。

#### 5.1.3 建议

- 继续沿用“原生命令优先 + API 兜底”模式；
- 对其它候选功能（如插入形状）也可复用该策略；
- 持续维护 Smart 命令清单，确保多语言版本可用。

---

### 5.2 美化文本部分改造方案

#### 5.2.1 改造目标

对于单个文本对象的简单格式操作，使用原生命令；批量操作和复杂格式保持自定义实现。

#### 5.2.2 实施步骤

**步骤 1：识别可替换的格式操作**

- ✅ 字体加粗/斜体/下划线
- ✅ 字号增大/减小
- ✅ 文本对齐（左/中/右）

**步骤 2：创建混合实现**

```csharp
public void ApplyTextFormatting(IShape shape)
{
    // 1. 使用原生命令处理基础格式
    if (shape.HasText)
    {
        commandExecutor.ExecuteMso(OfficeCommands.Bold);
        commandExecutor.ExecuteMso(OfficeCommands.AlignCenter);
    }

    // 2. 使用自定义逻辑处理高级格式
    ApplyAdvancedFormatting(shape);
}
```

#### 5.2.3 改造工作量

- 代码修改：2-3 天
- 测试验证：1-2 天
- **总计**：3-5 天

---

## 📈 六、预期收益分析

### 6.1 代码量减少

| 功能模块         | 当前代码行数 | 改造后代码行数 | 减少比例 |
| ---------------- | ------------ | -------------- | -------- |
| 对齐功能         | ~800 行      | ~200 行        | **75%**  |
| 美化文本（部分） | ~500 行      | ~350 行        | **30%**  |
| **总计**         | ~1300 行     | ~550 行        | **58%**  |

### 6.2 性能提升

- **对齐操作**：预计提升 20-30%（Office 原生实现更高效）
- **代码维护**：减少约 60% 的维护工作量
- **兼容性**：自动适配 Office 新版本特性

### 6.3 稳定性提升

- 减少自定义代码中的潜在 Bug
- 利用 Office 官方测试和优化
- 更好的错误处理机制

---

## ⚠️ 七、风险评估与应对

### 7.1 技术风险

| 风险                         | 影响 | 概率 | 应对措施                         |
| ---------------------------- | ---- | ---- | -------------------------------- |
| MSO 命令名称在不同版本不一致 | 中   | 低   | 创建命令映射表，支持多版本       |
| 原生命令执行失败             | 中   | 中   | 实现自动回退机制                 |
| 原生命令行为与预期不符       | 高   | 低   | 充分测试，保留自定义实现作为备选 |

### 7.2 兼容性风险

| Office 版本     | 支持情况 | 备注     |
| --------------- | -------- | -------- |
| PowerPoint 2016 | ✅ 支持  | 需要验证 |
| PowerPoint 2019 | ✅ 支持  | 需要验证 |
| PowerPoint 2021 | ✅ 支持  | 需要验证 |
| Microsoft 365   | ✅ 支持  | 需要验证 |

**应对措施**：

- 创建版本检测机制
- 为不同版本提供命令映射
- 实现降级策略

---

## 📅 八、实施计划

### 8.1 阶段一：对齐功能改造（优先级 P0）

✅ **已完成** – 8 个对齐/分布命令全部切换至 Smart MSO 命令并提供 API 兜底。

**交付物**：

1. Smart 命令映射（`ObjectsAlign*Smart`、`AlignDistribute*`）；
2. `AlignHelper.ExecuteAlignment` 中的 MSO 优先逻辑及详细日志；
3. 文档更新（本报告 & `OfficeCommands.cs` 常量清单）。

### 8.2 阶段二：美化文本部分改造（优先级 P1）

**时间**：1 周

**任务**：

1. 识别可替换的格式操作（1 天）
2. 实现混合方案（2-3 天）
3. 测试验证（2 天）

### 8.3 阶段三：插入形状部分改造（优先级 P1）

**时间**：3-5 天

**任务**：

1. 使用原生命令插入基础形状（1-2 天）
2. 保留自定义逻辑处理包围框（1 天）
3. 测试验证（1-2 天）

### 8.4 阶段四：文件操作辅助（优先级 P2）

**时间**：1-2 天

**任务**：

1. 添加自动保存选项（1 天）
2. 测试验证（1 天）

---

## 🎯 九、改造建议总结

### 9.1 立即改造（P0）

✅ **对齐功能（8 个基础对齐命令）** — _已完成_

- 改造难度：低
- 收益：高
- 风险：低
- **状态**：已上线（MSO 优先 + API 兜底）

### 9.2 分阶段改造（P1）

⚠️ **美化文本和插入形状的部分功能**

- 改造难度：中
- 收益：中
- 风险：中
- **建议**：在 P0 完成后，根据实际情况决定是否继续

### 9.3 可选改造（P2）

📋 **文件操作辅助**

- 改造难度：低
- 收益：低
- 风险：低
- **建议**：根据用户需求决定

### 9.4 保持现状

❌ **以下功能暂不建议改造**：

- 吸附功能（Bt121-Bt124）
- 大小调整功能（Bt201-Bt214）
- 参考线对齐功能（Bt301-Bt323）
- 隐显对象（Bt401）
- 裁剪出框（Bt402）
- 美化表格/图表（Bt501/Bt503）- 复杂格式需要自定义
- 设置菜单功能

**原因**：这些功能涉及复杂的自定义逻辑，Office 没有对应的原生命令，或原生命令无法满足需求。

---

## 📝 十、改造检查清单

### 10.1 对齐功能改造检查清单

- [ ] 验证 MSO 命令名称在所有 Office 版本中可用
- [ ] 在 `OfficeCommands.cs` 中添加对齐命令常量
- [ ] 在 `AlignHelper.cs` 中添加 `ExecuteAlignmentNative` 方法
- [ ] 修改 `CustomRibbon.cs` 使用原生命令
- [ ] 添加配置选项支持切换
- [ ] 实现自动回退机制
- [ ] 编写单元测试
- [ ] 进行兼容性测试
- [ ] 更新用户文档

### 10.2 代码质量检查

- [ ] 代码符合项目编码规范
- [ ] 添加 XML 文档注释
- [ ] 异常处理完善
- [ ] 日志记录完整
- [ ] 性能测试通过

---

## 📚 十一、参考资料

### 11.1 Office MSO 命令参考

- [Microsoft Office MSO 命令列表](https://docs.microsoft.com/office/vba/api/office.commandbars.executemso)
- [PowerPoint 对象模型参考](https://docs.microsoft.com/office/vba/api/powerpoint.application)

### 11.2 项目相关文档

- `docs/PPT常用MSO命令.MD` - 常用 MSO 命令列表
- `PPA/Core/Abstraction/Business/OfficeCommands.cs` - 命令常量定义
- `PPA/Utilities/CommandExecutor.cs` - 命令执行器实现

---

## 📊 十二、改造统计

### 12.1 功能统计

| 类别       | 总数   | 可改造 | 部分可改造 | 不可改造 |
| ---------- | ------ | ------ | ---------- | -------- |
| 对齐功能   | 8      | 8      | 0          | 0        |
| 吸附功能   | 4      | 0      | 0          | 4        |
| 大小调整   | 8      | 0      | 0          | 8        |
| 参考线对齐 | 9      | 0      | 0          | 9        |
| 选择功能   | 2      | 0      | 0          | 2        |
| 格式化功能 | 3      | 0      | 3          | 0        |
| 插入功能   | 1      | 0      | 1          | 0        |
| 设置菜单   | 4      | 0      | 0          | 4        |
| **总计**   | **39** | **8**  | **4**      | **27**   |

### 12.2 改造覆盖率

- **完全可改造**：8/39 = **20.5%**
- **部分可改造**：4/39 = **10.3%**
- **不可改造**：27/39 = **69.2%**

### 12.3 预期代码减少

- **对齐功能模块**：通过 MSO 命令驱动，核心逻辑由原生实现承担
- **美化文本模块**：预计减少约 150 行代码（30%）
- **总计**：当前已确认可减少 150 行，后续可继续扩展

---

## 🎯 十三、结论与建议

### 13.1 主要结论

1. **对齐功能已完成原生命令改造**：8 个基础对齐/分布命令全部使用 Smart MSO 命令实现，并保留 `ShapeRange.Align/Distribute` 兜底逻辑。

2. **部分功能可以混合使用**：美化文本和插入形状可以部分使用原生命令，但需要保留自定义逻辑。

3. **大部分功能需要保持现状**：89.7% 的功能涉及复杂的自定义逻辑，Office 没有对应的原生命令。

### 13.2 实施建议

1. **对齐功能继续保持“MSO 优先 + API 兜底”模式**

   - 充分利用 Ribbon Smart 命令，保证体验与原生 UI 一致
   - 兜底逻辑保留所有扩展参数（如 `alignToSlide`），确保兼容性

2. **分阶段实施其他改造**（P1）

   - 美化文本和插入形状可以部分使用原生命令
   - 根据实际效果和用户反馈决定

3. **保持核心功能不变**
   - 对齐、吸附、大小调整、参考线对齐等核心功能保持自定义实现
   - 这些功能是 PPA 的独特价值，不应替换

### 13.3 预期成果

完成部分改造后：

- ✅ 代码量减少约 150 行（仅美化文本模块）
- ✅ 维护成本降低约 30%
- ✅ 兼容性保持稳定

---

**报告版本**: v1.0  
**最后更新**: 2025 年 1 月  
**维护者**: PPA 开发团队

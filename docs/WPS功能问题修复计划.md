# WPS 功能问题修复计划

## 当前状态

✅ **插件已成功在 WPS 中加载**

- 注册表配置正确
- 插件可以正常启动
- 大部分功能正常工作

## 发现的问题

### 1. 参考线对齐功能不生效

**原因分析：**

- `GuideAlign*` 方法直接访问 `app.ActivePresentation.Guides`，这是 PowerPoint 特有的 API
- WPS 可能不支持 `Guides` 集合，或 API 不同
- `InvokeWithNative` 尝试将 WPS 的 `IApplication` 转换为 `NETOP.Application`，转换失败

**解决方案：**

- 检查 WPS 是否支持参考线 API
- 如果不支持，实现基于 WPS API 的替代方案
- 或提示用户该功能在 WPS 中不可用

### 2. 美化表格、美化文本、美化图表不生效

**原因分析：**

- `FormatTables(ITable)` 和 `ApplyTextFormatting(IShape)` 尝试将 WPS 的 COM 对象转换为 `NETOP.Table` 和 `NETOP.Shape`
- WPS 的 COM 对象是 `dynamic` 类型，不是 NetOffice 类型，转换失败
- 代码直接返回，没有执行实际的美化操作

**解决方案：**

- 在 WPS 适配器中实现格式化功能
- 或通过 `dynamic` 调用 WPS 的 COM API 实现相同的格式化逻辑
- 需要为 `WpsTable`、`WpsTextRange`、`WpsChart` 添加格式化方法

### 3. 裁剪出框可能删除原图形

**原因分析：**

- `MSOICrop` 可能使用了 PowerPoint 特有的 API
- WPS 的布尔运算 API 可能不同

**解决方案：**

- 检查 `MSOICrop` 的实现
- 适配 WPS 的布尔运算 API
- 添加错误处理和回滚机制

## 修复优先级

1. **高优先级**：美化表格/文本/图表（核心功能）
2. **中优先级**：参考线对齐（辅助功能）
3. **低优先级**：裁剪出框（边缘功能）

## 实施步骤

### 步骤 1：修复美化功能

1. 为 `WpsTable` 添加格式化方法
2. 为 `WpsTextRange` 添加格式化方法
3. 为 `WpsChart` 添加格式化方法
4. 更新 `TableFormatHelper`、`TextFormatHelper`、`ChartFormatHelper` 以支持 WPS

### 步骤 2：修复参考线对齐

1. 检查 WPS 是否支持参考线 API
2. 如果不支持，实现替代方案或提示用户

### 步骤 3：修复裁剪功能

1. 检查 `MSOICrop` 的实现
2. 适配 WPS 的布尔运算 API

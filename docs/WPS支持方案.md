# PPA 项目 WPS Office 支持方案

**版本**: v1.0  
**制定日期**: 2025 年 1 月  
**状态**: 📋 方案设计阶段

---

## 📋 一、项目现状分析

### 1.1 当前架构

PPA 项目当前基于以下技术栈：

- **开发框架**: VSTO (Visual Studio Tools for Office)
- **API 包装**: NetOffice.PowerPointApi
- **原生接口**: Microsoft.Office.Interop.PowerPoint
- **插件类型**: COM 加载项 (Add-in)
- **目标平台**: Microsoft PowerPoint 2016+

### 1.2 核心依赖

```
PPA 项目
├── VSTO Runtime (Microsoft.Office.Tools)
├── NetOffice.PowerPointApi
├── Microsoft.Office.Interop.PowerPoint
└── COM 互操作层
```

### 1.3 主要功能模块

1. **格式化模块** (Formatting/)

   - 表格格式化 (TableFormatHelper)
   - 文本格式化 (TextFormatHelper)
   - 图表格式化 (ChartFormatHelper)

2. **UI 模块** (UI/)

   - 自定义 Ribbon (CustomRibbon)
   - 快捷键管理 (KeyboardShortcutHelper)
   - 设置对话框 (SettingsForm)

3. **核心模块** (Core/)
   - 异常处理 (ExHandler)
   - 资源管理 (ResourceManager)
   - 性能监控 (Profiler)

---

## 🎯 二、WPS 支持目标

### 2.1 支持范围

- ✅ **WPS 演示 (WPS Presentation)** - 主要目标
- ⚠️ **WPS 表格 (WPS Spreadsheet)** - 可选（如果需求）
- ❌ **WPS 文字 (WPS Writer)** - 暂不支持

### 2.2 功能兼容性目标

| 功能模块   | PowerPoint | WPS 演示    | 优先级 |
| ---------- | ---------- | ----------- | ------ |
| 表格格式化 | ✅         | 🎯 目标支持 | 高     |
| 文本格式化 | ✅         | 🎯 目标支持 | 高     |
| 图表格式化 | ✅         | 🎯 目标支持 | 中     |
| 对齐工具   | ✅         | 🎯 目标支持 | 高     |
| 批量操作   | ✅         | 🎯 目标支持 | 高     |
| Ribbon UI  | ✅         | ⚠️ 需适配   | 高     |
| 快捷键     | ✅         | ⚠️ 需适配   | 中     |
| 配置文件   | ✅         | ✅ 通用     | 高     |

---

## 🔍 三、技术挑战分析

### 3.1 WPS 与 PowerPoint 的差异

#### 3.1.1 COM API 兼容性

**PowerPoint COM API**:

- 完整的 COM 接口支持
- 标准的 Office 对象模型
- 完善的类型库

**WPS COM API**:

- ⚠️ **部分兼容** Microsoft Office COM API
- ⚠️ 某些接口可能缺失或不完全兼容
- ⚠️ 类型库可能不同
- ✅ 支持基本的 COM 互操作

#### 3.1.2 插件架构差异

**PowerPoint (VSTO)**:

```
VSTO Runtime → PowerPoint Add-in → COM 加载项
```

**WPS 演示**:

```
WPS Add-in 平台 → JavaScript/C# 插件 → COM 互操作
```

#### 3.1.3 功能差异

| 功能      | PowerPoint  | WPS 演示    | 影响 |
| --------- | ----------- | ----------- | ---- |
| 表格样式  | ✅ 完整支持 | ⚠️ 部分支持 | 中等 |
| 主题颜色  | ✅ 完整支持 | ⚠️ 可能不同 | 中等 |
| Ribbon UI | ✅ 完整支持 | ⚠️ 需验证   | 高   |
| 撤销/重做 | ✅ 完整支持 | ✅ 支持     | 低   |
| 形状操作  | ✅ 完整支持 | ✅ 基本支持 | 低   |

### 3.2 主要技术挑战

1. **VSTO 依赖**: WPS 不支持 VSTO Runtime
2. **NetOffice 兼容性**: 需要验证 NetOffice 是否支持 WPS
3. **Ribbon UI**: WPS 的 Ribbon 实现可能不同
4. **API 差异**: 某些 PowerPoint 特有 API 在 WPS 中可能不存在
5. **插件加载机制**: WPS 的插件加载方式与 PowerPoint 不同

---

## 🏗️ 四、技术方案设计

### 4.1 方案选择

经过分析，推荐采用 **"抽象层 + 平台适配器"** 架构：

```
┌─────────────────────────────────────┐
│         PPA 业务逻辑层              │
│  (Formatting, UI, Core 等模块)      │
└─────────────────────────────────────┘
              ↓
┌─────────────────────────────────────┐
│      抽象接口层 (IApplication)      │
│  - IApplication                     │
│  - IPresentation                    │
│  - ITable, ITextRange, IChart 等     │
└─────────────────────────────────────┘
              ↓
┌──────────────────┬──────────────────┐
│  PowerPoint 适配器│   WPS 适配器      │
│  - PPTApplication │  - WPSApplication │
│  - PPTTable      │  - WPSTable       │
│  - PPTTextRange  │  - WPSTextRange   │
└──────────────────┴──────────────────┘
              ↓
┌──────────────────┬──────────────────┐
│  NetOffice API   │   WPS COM API     │
│  PowerPointApi   │   (直接 COM)       │
└──────────────────┴──────────────────┘
```

### 4.2 架构设计

#### 4.2.1 抽象接口层

创建平台无关的接口定义：

```csharp
// Core/Abstraction/IApplication.cs
namespace PPA.Core.Abstraction
{
    /// <summary>
    /// 应用程序抽象接口（支持 PowerPoint 和 WPS）
    /// </summary>
    public interface IApplication
    {
        ApplicationType Type { get; }
        IPresentation ActivePresentation { get; }
        // ... 其他通用接口
    }

    public enum ApplicationType
    {
        PowerPoint,
        WPSPresentation
    }
}
```

#### 4.2.2 平台适配器层

为每个平台创建适配器：

```csharp
// Core/Adapters/PowerPoint/PowerPointApplication.cs
namespace PPA.Core.Adapters.PowerPoint
{
    public class PowerPointApplication : IApplication
    {
        private readonly NETOP.Application _netApp;
        // 实现 IApplication 接口
    }
}

// Core/Adapters/WPS/WPSApplication.cs
namespace PPA.Core.Adapters.WPS
{
    public class WPSApplication : IApplication
    {
        private dynamic _wpsApp; // WPS COM 对象
        // 实现 IApplication 接口
    }
}
```

#### 4.2.3 工厂模式

使用工厂模式创建平台适配器：

```csharp
// Core/Adapters/ApplicationFactory.cs
namespace PPA.Core.Adapters
{
    public static class ApplicationFactory
    {
        public static IApplication Create(object nativeApp)
        {
            // 检测应用程序类型
            if (IsPowerPoint(nativeApp))
                return new PowerPointApplication(nativeApp);
            else if (IsWPS(nativeApp))
                return new WPSApplication(nativeApp);
            else
                throw new NotSupportedException("不支持的应用程序类型");
        }

        private static bool IsPowerPoint(object app) { /* ... */ }
        private static bool IsWPS(object app) { /* ... */ }
        private static ApplicationType DetectApplicationType() { /* ... */ }
    }
}
```

---

## 📦 五、实施计划

### 5.1 阶段一：架构重构（2-3 周）

#### 5.1.1 创建抽象层

- [ ] 创建 `PPA.Core.Abstraction` 命名空间
- [ ] 定义核心接口：
  - `IApplication`
  - `IPresentation`
  - `ISlide`
  - `ITable`
  - `ITextRange`
  - `IChart`
  - `IShape`
- [ ] 定义枚举类型：
  - `ApplicationType`
  - `FeatureSupportLevel`

#### 5.1.2 创建 PowerPoint 适配器

- [ ] 创建 `PPA.Core.Adapters.PowerPoint` 命名空间
- [ ] 实现 PowerPoint 适配器类：
  - `PowerPointApplication`
  - `PowerPointPresentation`
  - `PowerPointTable`
  - `PowerPointTextRange`
  - `PowerPointChart`
- [ ] 将现有代码迁移到适配器模式

#### 5.1.3 重构现有代码

- [ ] 修改 `ThisAddIn.cs` 使用 `IApplication`
- [ ] 修改 `FormatHelper` 系列使用抽象接口
- [ ] 修改 `CustomRibbon` 使用抽象接口
- [ ] 保持向后兼容性

### 5.2 阶段二：WPS 适配器开发（3-4 周）

#### 5.2.1 WPS COM API 研究

- [ ] 研究 WPS 演示 COM API 文档
- [ ] 创建 WPS 测试环境
- [ ] 验证 WPS COM 互操作性
- [ ] 识别 API 差异和限制

#### 5.2.2 实现 WPS 适配器

- [ ] 创建 `PPA.Core.Adapters.WPS` 命名空间
- [ ] 实现 WPS 适配器类：
  - `WPSApplication`
  - `WPSPresentation`
  - `WPSTable`
  - `WPSTextRange`
  - `WPSChart`
- [ ] 处理 API 差异和兼容性问题

#### 5.2.3 WPS 插件加载机制

- [ ] 研究 WPS 插件开发文档
- [ ] 创建 WPS 插件入口点
- [ ] 实现 WPS 插件注册机制
- [ ] 处理 WPS 特定的初始化逻辑

### 5.3 阶段三：功能适配（2-3 周）

#### 5.3.1 核心功能适配

- [ ] 表格格式化功能适配
- [ ] 文本格式化功能适配
- [ ] 图表格式化功能适配
- [ ] 对齐工具适配

#### 5.3.2 UI 适配

- [ ] Ribbon UI 适配（如果 WPS 支持）
- [ ] 快捷键系统适配
- [ ] 对话框和设置界面适配

#### 5.3.3 功能降级处理

- [ ] 识别 WPS 不支持的功能
- [ ] 实现功能检测机制
- [ ] 提供友好的降级提示

### 5.4 阶段四：测试和优化（2 周）

#### 5.4.1 功能测试

- [ ] PowerPoint 功能回归测试
- [ ] WPS 功能测试
- [ ] 跨平台兼容性测试

#### 5.4.2 性能优化

- [ ] 性能对比测试
- [ ] 优化 WPS 适配器性能
- [ ] 内存泄漏检查

#### 5.4.3 文档更新

- [ ] 更新 README.md
- [ ] 更新部署指南
- [ ] 创建 WPS 安装说明

---

## 🛠️ 六、技术实现细节

### 6.1 应用程序类型检测

```csharp
public static class ApplicationDetector
{
    public static ApplicationType DetectType(object app)
    {
        if (app == null) return ApplicationType.Unknown;

        string progId = app.GetType().Name;
        string typeName = app.GetType().FullName;

        // 检测 PowerPoint
        if (typeName.Contains("Microsoft.Office.Interop.PowerPoint") ||
            typeName.Contains("NetOffice.PowerPointApi"))
        {
            return ApplicationType.PowerPoint;
        }

        // 检测 WPS
        if (typeName.Contains("WPS") ||
            progId.Contains("WPS") ||
            IsWPSApplication(app))
        {
            return ApplicationType.WPSPresentation;
        }

        return ApplicationType.Unknown;
    }

    private static bool IsWPSApplication(object app)
    {
        try
        {
            // 尝试访问 WPS 特有属性
            dynamic dynApp = app;
            string name = dynApp.Name;
            return name != null && name.Contains("WPS");
        }
        catch
        {
            return false;
        }
    }
}
```

### 6.2 WPS COM 对象访问

```csharp
public class WPSApplication : IApplication
{
    private dynamic _wpsApp;

    public WPSApplication(object nativeApp)
    {
        _wpsApp = nativeApp;
        // 验证 WPS 对象
        ValidateWPSObject();
    }

    private void ValidateWPSObject()
    {
        try
        {
            // 尝试访问 WPS 基本属性
            string name = _wpsApp.Name;
            if (!name.Contains("WPS"))
                throw new InvalidOperationException("不是有效的 WPS 应用程序对象");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"WPS 对象验证失败: {ex.Message}", ex);
        }
    }

    public IPresentation ActivePresentation
    {
        get
        {
            try
            {
                dynamic activePres = _wpsApp.ActivePresentation;
                return new WPSPresentation(activePres);
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"获取 WPS 活动演示文稿失败: {ex.Message}");
                return null;
            }
        }
    }
}
```

### 6.3 功能兼容性检查

```csharp
public class FeatureCompatibility
{
    public static bool IsFeatureSupported(IApplication app, Feature feature)
    {
        switch (app.Type)
        {
            case ApplicationType.PowerPoint:
                return true; // PowerPoint 支持所有功能

            case ApplicationType.WPSPresentation:
                return CheckWPSFeatureSupport(app, feature);

            default:
                return false;
        }
    }

    private static bool CheckWPSFeatureSupport(IApplication app, Feature feature)
    {
        // WPS 功能支持检查表
        var wpsSupport = new Dictionary<Feature, bool>
        {
            { Feature.TableFormatting, true },
            { Feature.TextFormatting, true },
            { Feature.ChartFormatting, true },
            { Feature.AdvancedTableStyles, false }, // WPS 可能不支持高级表格样式
            { Feature.ThemeColors, true },
            // ... 其他功能
        };

        return wpsSupport.GetValueOrDefault(feature, false);
    }
}
```

### 6.4 错误处理和降级

```csharp
public class SafeOperation
{
    public static TResult ExecuteWithFallback<TResult>(
        Func<TResult> primaryAction,
        Func<TResult> fallbackAction,
        string operationName)
    {
        try
        {
            return primaryAction();
        }
        catch (COMException ex)
        {
            Profiler.LogMessage($"{operationName} 主操作失败，尝试降级方案: {ex.Message}");
            try
            {
                return fallbackAction();
            }
            catch (Exception fallbackEx)
            {
                Profiler.LogMessage($"{operationName} 降级方案也失败: {fallbackEx.Message}");
                throw;
            }
        }
    }
}
```

---

## 📋 七、项目结构调整

### 7.1 新增目录结构

```
PPA/
├── Core/
│   ├── Abstraction/          # 新增：抽象接口层
│   │   ├── IApplication.cs
│   │   ├── IPresentation.cs
│   │   ├── ITable.cs
│   │   ├── ITextRange.cs
│   │   └── ...
│   │
│   └── Adapters/             # 新增：平台适配器层
│       ├── PowerPoint/       # PowerPoint 适配器
│       │   ├── PowerPointApplication.cs
│       │   ├── PowerPointTable.cs
│       │   └── ...
│       │
│       ├── WPS/              # WPS 适配器
│       │   ├── WPSApplication.cs
│       │   ├── WPSTable.cs
│       │   └── ...
│       │
│       └── ApplicationFactory.cs
│
├── Formatting/                # 修改：使用抽象接口
│   ├── TableFormatHelper.cs  # 改为使用 ITable
│   ├── TextFormatHelper.cs   # 改为使用 ITextRange
│   └── ...
│
└── UI/                       # 修改：使用抽象接口
    └── CustomRibbon.cs       # 改为使用 IApplication
```

### 7.2 命名空间规划

```csharp
PPA.Core.Abstraction          // 抽象接口
PPA.Core.Adapters             // 适配器基类
PPA.Core.Adapters.PowerPoint  // PowerPoint 适配器
PPA.Core.Adapters.WPS         // WPS 适配器
PPA.Core.Compatibility        // 兼容性检查
```

---

## ⚠️ 八、风险和限制

### 8.1 技术风险

1. **WPS COM API 兼容性未知**

   - **风险**: WPS COM API 可能与 PowerPoint 不完全兼容
   - **缓解**: 充分测试，实现降级方案

2. **NetOffice 不支持 WPS**

   - **风险**: NetOffice 可能不支持 WPS
   - **缓解**: 直接使用 WPS COM API，不依赖 NetOffice

3. **Ribbon UI 兼容性**

   - **风险**: WPS 的 Ribbon 实现可能不同
   - **缓解**: 研究 WPS UI 扩展机制，必要时使用替代方案

4. **性能影响**
   - **风险**: 抽象层可能带来性能开销
   - **缓解**: 优化适配器实现，最小化性能损失

### 8.2 功能限制

1. **部分 PowerPoint 特有功能可能无法在 WPS 中实现**
2. **WPS 的某些 API 行为可能与 PowerPoint 不同**
3. **需要维护两套适配器代码**

### 8.3 维护成本

- 需要同时维护 PowerPoint 和 WPS 两套适配器
- 需要持续关注 WPS API 更新
- 需要处理平台差异带来的 bug

---

## 📊 九、可行性评估

### 9.1 技术可行性

| 方面       | 可行性 | 说明                    |
| ---------- | ------ | ----------------------- |
| COM 互操作 | ✅ 高  | WPS 支持 COM 互操作     |
| API 兼容性 | ⚠️ 中  | 需要验证具体 API 兼容性 |
| 插件架构   | ⚠️ 中  | 需要研究 WPS 插件机制   |
| 抽象层设计 | ✅ 高  | 标准设计模式，可行      |
| 性能影响   | ✅ 高  | 抽象层开销可控          |

### 9.2 工作量估算

- **架构重构**: 2-3 周
- **WPS 适配器开发**: 3-4 周
- **功能适配**: 2-3 周
- **测试和优化**: 2 周
- **总计**: **9-12 周**（约 2-3 个月）

### 9.3 资源需求

- 开发人员：1-2 人
- WPS 测试环境
- WPS 开发文档
- 测试用例和测试数据

---

## 🎯 十、建议和决策

### 10.1 推荐方案

**采用 "抽象层 + 平台适配器" 架构**，理由：

1. ✅ **可维护性**: 清晰的架构，易于维护和扩展
2. ✅ **可扩展性**: 未来可以轻松支持其他 Office 软件
3. ✅ **向后兼容**: 不影响现有 PowerPoint 功能
4. ✅ **代码复用**: 业务逻辑层可以复用

### 10.2 实施建议

1. **分阶段实施**: 按照计划分阶段进行，降低风险
2. **充分测试**: 每个阶段都要充分测试
3. **文档先行**: 先完成架构设计，再开始编码
4. **原型验证**: 先实现一个简单功能验证可行性

### 10.3 备选方案

如果 WPS COM API 兼容性太差，可以考虑：

1. **独立 WPS 版本**: 为 WPS 开发独立版本
2. **功能子集**: 只支持 WPS 中兼容的功能
3. **Web 插件**: 考虑 WPS 的 Web 插件方案（如果适用）

---

## 📝 十一、下一步行动

### 11.1 立即行动

1. [ ] **验证 WPS COM API 可用性**

   - 安装 WPS 演示
   - 创建简单的 COM 互操作测试程序
   - 验证基本 API 调用

2. [ ] **研究 WPS 插件开发文档**

   - 查找 WPS 官方文档
   - 了解 WPS 插件开发流程
   - 了解 WPS 插件注册机制

3. [ ] **创建技术验证原型**
   - 实现简单的 WPS 适配器原型
   - 验证核心 API 调用
   - 评估技术可行性

### 11.2 短期计划（1-2 周）

1. [ ] 完成架构设计评审
2. [ ] 创建抽象接口定义
3. [ ] 实现 PowerPoint 适配器（重构现有代码）
4. [ ] 完成阶段一验收

### 11.3 中期计划（1-2 月）

1. [ ] 完成 WPS 适配器开发
2. [ ] 完成功能适配
3. [ ] 完成初步测试

---

## 📚 十二、参考资料

### 12.1 WPS 开发资源

- WPS 开放平台: https://open.wps.cn/
- WPS 插件开发文档（需要查找）
- WPS COM API 文档（需要查找）

### 12.2 技术文档

- COM 互操作最佳实践
- 设计模式：适配器模式、工厂模式
- .NET Framework COM 互操作

---

## ✅ 总结

本方案提出了通过 **"抽象层 + 平台适配器"** 架构来支持 WPS Office，具有以下优势：

1. ✅ **架构清晰**: 分离关注点，易于维护
2. ✅ **向后兼容**: 不影响现有 PowerPoint 功能
3. ✅ **可扩展**: 未来可以支持更多平台
4. ✅ **风险可控**: 分阶段实施，降低风险

**建议**: 先进行技术验证，确认 WPS COM API 的可用性和兼容性，再开始全面实施。

---

**文档版本**: v1.0  
**最后更新**: 2025 年 1 月  
**维护者**: PPA 开发团队

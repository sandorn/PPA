# PPA - PowerPoint 增强插件

**版本**: v1.2.0  
**发布日期**: 2025 年 1 月

PPA 是一个功能强大的 PowerPoint 增强插件，提供多种实用工具来提升 PowerPoint 演示文稿的编辑效率和质量。

## 项目介绍

PPA（PowerPoint Advanced Add-in）是基于 .NET Framework 开发的 PowerPoint 插件，使用 NetOffice 库实现与 PowerPoint 的交互。该插件提供自定义功能区，包含多种实用工具，帮助用户更高效地创建和编辑演示文稿。

## 主要功能

### 核心功能

- **自定义功能区集成** - 无缝集成到 PowerPoint 界面
- **图形对齐和格式化工具** - 快速对齐、分布、吸附形状
- **批量操作功能** - 批量格式化表格、文本、图表
- **扩展格式化选项** - 可配置的格式化样式
- **形状处理工具** - 形状创建、裁剪、显示/隐藏
- **界面交互提示** - Toast 通知和进度指示

### 高级特性

- **异步操作支持** - 耗时操作异步执行，不阻塞 UI
- **配置化支持** - XML 配置文件自定义格式化样式
- **多语言支持** - 支持中文（简体）和英文界面
- **撤销优化** - 统一的撤销/重做管理
- **快捷键支持** - 全局快捷键快速执行常用操作
- **设置菜单** - 语言切换、参数配置、关于信息

## 技术栈

- **C#** - 主要开发语言
- **.NET Framework 4.8** - 目标框架
- **NetOffice** - PowerPoint API 包装库（主要使用）
- **Microsoft Office Interop** - 原生 Office COM 互操作（特定场景）
- **Microsoft.Extensions.DependencyInjection** - 依赖注入容器
- **Windows Forms** - UI 对话框支持
- **Windows API** - 全局快捷键支持

## 项目结构

```
PPA/
├── Core/                     # 核心基础设施模块
│   ├── ExHandler.cs         # 异常处理类
│   ├── Profiler.cs          # 性能分析工具
│   ├── ResourceManager.cs   # 多语言资源管理
│   ├── ApplicationProvider.cs # 应用程序提供者
│   ├── Abstraction/         # 抽象接口层
│   │   ├── Business/        # 业务接口
│   │   ├── Infrastructure/  # 基础设施接口
│   │   └── Presentation/    # 表示层接口
│   ├── Adapters/            # 适配器实现
│   │   └── PowerPoint/      # PowerPoint适配器
│   ├── DI/                  # 依赖注入配置
│   │   └── ServiceCollectionExtensions.cs
│   └── Logging/             # 日志系统
│       ├── ILogger.cs       # 日志接口
│       └── ProfilerLoggerAdapter.cs # 日志适配器
│
├── Formatting/               # 格式化业务模块
│   ├── TableFormatHelper.cs # 表格格式化
│   ├── TextFormatHelper.cs  # 文本格式化
│   ├── ChartFormatHelper.cs # 图表格式化
│   ├── TableBatchHelper.cs  # 表格批量操作
│   ├── TextBatchHelper.cs   # 文本批量操作
│   ├── ChartBatchHelper.cs  # 图表批量操作
│   ├── ShapeBatchHelper.cs  # 形状批量操作
│   ├── FormattingConfig.cs  # 格式化配置管理
│   ├── UndoHelper.cs        # 撤销操作管理
│   └── AlignHelper.cs       # 对齐辅助工具
│
├── Shape/                    # 形状处理模块
│   ├── ShapeUtils.cs        # 形状处理工具
│   └── MSOICrop.cs          # Microsoft Office交互裁剪类
│
├── UI/                       # UI模块
│   ├── CustomRibbon.cs      # 自定义功能区实现
│   ├── Providers/           # 提供者实现
│   │   ├── RibbonXmlProvider.cs    # Ribbon XML提供者
│   │   ├── RibbonIconProvider.cs   # Ribbon图标提供者
│   │   └── RibbonCommandRouter.cs  # Ribbon命令路由
│   ├── KeyboardShortcutHelper.cs # 全局快捷键管理
│   ├── Forms/               # 对话框窗体
│   │   ├── SettingsForm.cs # 设置对话框
│   │   └── AboutForm.cs    # 关于对话框
│   └── Ribbon.xml           # Ribbon配置文件
│
├── Utilities/                # 通用工具模块
│   ├── Toast.cs             # 通知提示类
│   ├── FileLocator.cs       # 文件定位工具
│   ├── ApplicationHelper.cs # 应用程序辅助类
│   └── AsyncOperationHelper.cs # 异步操作辅助类
│
├── Properties/               # 项目属性文件夹
│   ├── AssemblyInfo.cs
│   ├── Resources.resx        # 默认资源文件
│   ├── Resources.zh-CN.resx  # 中文资源文件
│   ├── Resources.en-US.resx # 英文资源文件
│   └── Settings.settings
│
├── Resources/                # 资源文件夹
│   └── icon/                # 图标资源
│
├── ThisAddIn.cs             # 插件主入口类
├── ThisAddIn.Designer.cs
├── ThisAddIn.Designer.xml
├── PPA.csproj               # 项目文件
├── PPA.sln                  # 解决方案文件
└── packages.config           # NuGet包配置
```

## 模块说明

### Core 模块

- **ExHandler.cs**: 统一异常处理类，提供异常捕获、日志记录和性能监控功能，支持 SafeGet/SafeSet 等安全访问方法
- **Profiler.cs**: 性能监控类，提供方法执行时间测量、记录和日志功能，支持文件日志和调试输出
- **ResourceManager.cs**: 多语言资源管理器，支持动态语言切换和本地化字符串管理
- **ApplicationProvider.cs**: 应用程序提供者，封装 Application 获取逻辑，消除静态依赖，提供 NetOffice 和 Native COM 对象访问
- **Abstraction/**: 抽象接口层，定义业务、基础设施和表示层接口，实现依赖倒置原则
  - **Business/**: 业务接口，如 ITextFormatHelper、IChartFormatHelper、ICommandExecutor 等
  - **Infrastructure/**: 基础设施接口，如 ILogger、IApplicationProvider 等
  - **Presentation/**: 表示层接口，如 IApplication、ISlide、IShape 等
- **Adapters/**: 适配器实现，将 NetOffice 对象适配为抽象接口
  - **AdapterUtils.cs**: 统一的适配器工具类，提供双向转换功能（Wrap/Unwrap）
  - **PowerPoint/**: PowerPoint 适配器实现，如 PowerPointApplication、PowerPointSlide 等
- **DI/**: 依赖注入配置，统一注册所有服务，使用 Microsoft.Extensions.DependencyInjection
- **Logging/**: 日志系统，提供统一的日志接口和实现
  - **ILogger.cs**: 日志接口定义，支持 LogInformation、LogWarning、LogError、LogDebug 等方法
  - **ProfilerLoggerAdapter.cs**: 基于 Profiler 的日志适配器实现
  - **LoggerProvider.cs**: 日志提供者，提供全局日志实例获取

### Utilities 模块

- **Toast.cs**: Toast 通知管理器，提供单消息框模式的用户提示
- **FileLocator.cs**: 文件定位工具，在多个可能的位置搜索文件
- **ApplicationHelper.cs**: 应用程序辅助类，提供 Application 对象转换和获取功能
- **AsyncOperationHelper.cs**: 异步操作辅助类，提供统一的异步操作执行框架，支持进度报告

### Formatting 模块

- **TableFormatHelper.cs**: 表格格式化辅助工具，提供表格样式格式化功能，支持表头、数据行、边框等样式配置
- **TextFormatHelper.cs**: 文本格式化辅助工具，提供文本样式格式化功能，支持字体、颜色、边距等配置
- **ChartFormatHelper.cs**: 图表格式化辅助工具，提供图表字体和样式格式化功能，支持标题、图例、坐标轴等格式化
- **TableBatchHelper.cs**: 表格批量操作类，提供批量格式化表格功能，支持选中表格或当前幻灯片所有表格
- **TextBatchHelper.cs**: 文本批量操作类，提供批量格式化文本功能，支持选中文本或当前幻灯片所有文本
- **ChartBatchHelper.cs**: 图表批量操作类，提供批量格式化图表功能，支持选中图表或当前幻灯片所有图表
- **ShapeBatchHelper.cs**: 形状批量操作类，提供形状的批量操作功能，如显示/隐藏、创建边界框等
- **FormattingConfig.cs**: 格式化配置管理类，从 XML 配置文件加载格式化参数
- **UndoHelper.cs**: 撤销操作管理类，统一管理 PowerPoint 的撤销/重做操作
- **AlignHelper.cs**: 对齐辅助工具类，提供形状对齐、分布、拉伸等操作

### Shape 模块

- **ShapeUtils.cs**: 形状处理工具类，实现 IShapeHelper 接口，提供形状创建、验证、选择等实用方法
- **MSOICrop.cs**: Microsoft Office 交互裁剪类，提供形状裁剪到幻灯片范围的功能，使用布尔运算实现精确裁剪

### UI 模块

- **CustomRibbon.cs**: 自定义功能区实现，处理 Ribbon UI 初始化和生命周期管理，集成 Ribbon XML、图标和命令路由服务
- **Providers/**: Ribbon 相关服务提供者
  - **EmbeddedRibbonXmlProvider.cs**: Ribbon XML 提供者，从嵌入式资源加载 Ribbon 配置文件
  - **RibbonIconProvider.cs**: Ribbon 图标提供者，负责管理和提供 Ribbon 图标，支持图标缓存
  - **RibbonCommandRouter.cs**: Ribbon 命令路由，负责处理按钮点击和命令执行，路由到相应的业务逻辑
- **KeyboardShortcutHelper.cs**: 全局快捷键管理，使用 Windows API 实现系统级快捷键，支持 Ctrl+1/2/3/4 等快捷键
- **Forms/**: Windows Forms 对话框
  - **SettingsForm.cs**: 设置对话框，用于编辑格式化配置和切换语言
  - **AboutForm.cs**: 关于对话框，显示插件版本和项目信息
- **Ribbon.xml**: Ribbon UI 配置文件，定义功能区布局和按钮，支持动态标签加载

## 安装说明

1. 确保已安装 PowerPoint（建议 2016 或更高版本）
2. 确保已安装 .NET Framework 4.8
3. 构建项目生成 DLL 文件
4. 将插件 DLL 注册到 PowerPoint
5. 重启 PowerPoint 后，插件将自动加载

## 使用方法

### 基本使用

安装完成后，在 PowerPoint 界面中会出现自定义功能区 "PPA 菜单"。点击相应的按钮即可使用各项功能。

### 主要功能说明

#### 对齐工具

- **左对齐/右对齐/顶对齐/底对齐** - 快速对齐选中的形状
- **水平居中/垂直居中** - 居中对齐形状
- **横向分布/纵向分布** - 均匀分布多个形状
- **吸附功能** - 快速吸附到参考线或页面边缘

#### 格式化工具

- **美化表格** - 批量格式化表格样式（支持异步执行，显示进度）
- **美化文本** - 批量格式化文本样式
- **美化图表** - 批量格式化图表字体和样式
  - 快捷键：`Ctrl+3`（全局快捷键，可在任何窗口使用）

#### 形状工具

- **插入形状** - 创建矩形外框或页面大小矩形
- **隐显对象** - 隐藏选中对象或显示所有隐藏对象
- **裁剪出框** - 将形状裁剪到幻灯片范围

#### 设置菜单

- **语言切换** - 在中文（简体）和英文之间切换界面语言
- **设置参数** - 编辑配置文件（`PPAConfig.xml`）
- **关于** - 查看插件版本和项目信息

### 配置文件

插件会在 `%AppData%\PPA\` 目录（通常为 `C:\Users\<用户名>\AppData\Roaming\PPA\`）创建 `PPAConfig.xml` 配置文件，用于自定义格式化样式和快捷键设置。首次运行时会自动生成默认配置文件。

配置文件支持自定义：

- 表格样式（边框、填充、字体等）
- 文本样式（字体、大小、颜色等）
- 图表样式（标题、图例、坐标轴字体等）
- 快捷键设置（美化表格、美化文本、美化图表、插入形状等功能的快捷键）
- 格式：只需配置数字或字母（如 `"3"`, `"C"`, `"F1"`），系统会自动添加 `Ctrl` 修饰键
- 示例：`FormatChart="3"` 表示 `Ctrl+3`，`FormatTables="T"` 表示 `Ctrl+T`

### 快捷键

- `Ctrl+3` - 美化图表（全局快捷键，可在 `PPAConfig.xml` 中自定义）
- 配置方式：在 `PPAConfig.xml` 的 `<Shortcuts>` 节点中，只需配置数字或字母（如 `FormatChart="3"`），系统会自动添加 `Ctrl` 修饰键
- 支持的键：数字 0-9、字母 A-Z、功能键 F1-F12

## 特性说明

### 异步操作支持

对于耗时操作（如批量格式化大量表格），插件采用异步执行方式，不会阻塞 PowerPoint UI。操作过程中会显示进度指示器，用户可以随时取消操作。

### 多语言支持

插件支持中文（简体）和英文两种语言，所有界面文本和提示信息都会根据用户选择的语言自动切换。语言设置保存在插件配置中，重启后仍然有效。

### 配置化支持

所有格式化样式都可以通过 XML 配置文件自定义，无需修改代码即可调整格式化行为。配置文件采用 XML 格式，易于编辑和理解。

### 撤销优化

所有操作都支持 PowerPoint 的撤销/重做功能，操作会被正确记录到撤销栈中，方便用户回退操作。

## 开发环境

- **Visual Studio 2019 或更高版本**
- **.NET Framework 4.8**
- **PowerPoint 2016 或更高版本**
- **VSTO (Visual Studio Tools for Office)**

### 依赖项

- NetOffice.PowerPointApi
- Microsoft.Office.Interop.PowerPoint
- Microsoft.Office.Tools.Common

## 项目结构说明

项目采用模块化设计，按功能划分为以下模块：

- **Core** - 核心基础设施（异常处理、性能监控、资源管理）
- **Utilities** - 通用工具（通知、文件定位、COM 对象扩展）
- **Formatting** - 格式化业务逻辑（表格、文本、图表格式化）
- **Shape** - 形状处理（形状创建、验证、裁剪）
- **UI** - 用户界面（Ribbon、快捷键、对话框）
- **AddIn** - 插件入口（初始化、生命周期管理）

## 开发指南

### 代码规范

- 遵循 C# 编码规范
- 使用 XML 文档注释
- 异常处理统一使用 `ExHandler.Run`
- 性能关键操作使用 `Profiler` 记录执行时间
- 用户提示使用 `Toast.Show`
- 所有用户可见文本使用 `ResourceManager.GetString` 进行本地化
- 日志记录统一使用 `ILogger` 接口（通过依赖注入获取）
- COM 对象生命周期管理统一使用 `using` 语句（NetOffice 对象已实现 `IDisposable`）
- 避免直接访问 `Globals.ThisAddIn`，通过 `IApplicationProvider` + `ApplicationHelper.GetNetOfficeApplication()` 获取应用上下文
- 如需访问原生 COM，必须置于 `NativeComGuard` 或等价封装中，并记录调用方、释放对象

### COM / NetOffice 使用规范

1. **入口层刷新策略**

   - Ribbon、快捷键、批处理等入口统一通过 `ApplicationHelper.EnsureValidNetApplication`（NetChannel）获取 `NETOP.Application`。
   - 批处理场景复用 `BatchContextHelper`（或同级封装）来重试 `ActiveWindow` / `Selection`。

2. **Native Guard**

   - 只有 Interop fallback 清单中的场景（例如全局命令执行、特定 Shape API）才可访问原生 COM。
   - 调用时必须通过 `ApplicationHelper.GetNativeComApplication(...)` 获得对象，并使用 `using` 语句确保自动释放。

3. **COM 对象释放**

   - 遍历 `CommandBars`、`ShapeRange`、`SlideRange` 等集合时使用 `using` 语句确保自动释放。
   - 禁止在字段级别缓存 RCW；只保留抽象接口或轻量标识。

4. **扫描与审查**
   - 提交或 CI 前执行 `tools/native-scan.ps1`（或 `rg "ApplicationHelper\.GetNativeComApplication" -g "*.cs"`）确保没有新增未受 Guard 保护的 native 访问。
   - 新增 native 调用时必须在 PR 描述中注明原因及 Guard 实现。

### 添加新功能

1. 在相应的模块目录下创建新的类文件
2. 在 `CustomRibbon.cs` 中添加按钮处理逻辑
3. 在 `Ribbon.xml` 中添加 UI 定义
4. 在资源文件中添加本地化文本
5. 更新 `README.md` 文档

### 测试

- 在 Visual Studio 中按 F5 启动调试
- PowerPoint 会自动启动并加载插件
- 检查日志输出以诊断问题

## 更新日志

### v1.2.0 (2025-01) - 架构完善版本

**主要变更**：

- ✅ 完成静态依赖消除，通过 `IApplicationProvider` 统一管理应用上下文
- ✅ 实现统一日志系统，所有模块使用 `ILogger` 接口
- ✅ 统一 COM 对象生命周期管理，使用 `using` 语句（NetOffice 对象已实现 `IDisposable`）
- ✅ 解耦 Ribbon 组件，分离为 `IRibbonXmlProvider`、`IRibbonIconProvider`、`IRibbonCommandRouter`
- ✅ 统一 NetOffice 与 Interop 使用场景，明确使用规范
- ✅ 优化美化文本、表格、图表功能，统一处理流程
- ✅ 改进日志输出，统一格式和级别

**技术改进**：

- 创建 `IApplicationProvider` 接口和实现，消除静态依赖
- 定义 `ILogger` 接口和 `LogLevel` 枚举，实现 `ProfilerLoggerAdapter`
- 统一使用 `using` 语句管理 COM 对象生命周期（NetOffice 对象已实现 `IDisposable`）
- 重构 Ribbon 组件，实现职责分离和依赖注入
- 统一格式化功能处理流程，避免 COM 对象生命周期问题
- 改进日志输出，使用统一的日志级别和格式

### v1.1.0 (2025-01) - 架构优化版本

**主要变更**：

- ✅ 完成依赖注入（DI）容器基础设施重构
- ✅ 实现平台抽象层架构，支持多平台扩展
- ✅ 将静态类重构为实例类，提升可测试性
- ✅ 创建 PowerPoint 适配器实现
- ✅ 重构业务逻辑使用 DI 和抽象接口
- ⚠️ 回退 WPS 适配支持，专注于 PowerPoint 平台优化
- ✅ 优化代码结构，提升可维护性

**技术改进**：

- 引入 `Microsoft.Extensions.DependencyInjection` DI 容器
- 创建平台抽象接口层（`IApplication`、`IShape`、`ITable` 等）
- 实现 PowerPoint 适配器（`PowerPointApplication`、`PowerPointShape` 等）
- 重构格式化辅助类支持依赖注入
- 优化代码组织，提升模块化程度

### v1.0.0 (2025-01) - 第一版正式发布

**主要特性**：

- ✅ 添加异步操作支持，提升大量表格格式化性能
- ✅ 实现多语言支持（中文/英文）
- ✅ 添加配置化支持，可通过 XML 自定义格式化样式
- ✅ 优化撤销/重做功能
- ✅ 添加全局快捷键支持（Ctrl+3 美化图表）
- ✅ 添加设置菜单（语言切换、参数配置、关于）
- ✅ 移除 VBA 依赖，提升性能和稳定性
- ✅ 重构项目结构，采用模块化设计
- ✅ 优化错误处理和日志记录
- ✅ 移除配置文件版本号管理，简化配置结构
- ✅ 移除旧路径配置文件迁移逻辑，统一使用新路径

## 贡献指南

欢迎提交 Issue 和 Pull Request 来帮助改进此项目。

### 贡献流程

1. Fork 本项目
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

### 代码提交规范

- 使用清晰的提交信息
-                确保代码通过编译和基本测试
- 更新相关文档

## 许可证

本项目采用 MIT 许可证。详见 LICENSE 文件。

## 项目链接

- **GitHub**: [https://github.com/sandorn/PPA](https://github.com/sandorn/PPA)
- **问题反馈**: [GitHub Issues](https://github.com/sandorn/PPA/issues)

## 致谢

感谢所有为这个项目做出贡献的开发者！

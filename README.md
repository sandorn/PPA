# PPA - PowerPoint 增强插件

**版本**: v1.0.0  
**发布日期**: 2025年1月

PPA 是一个功能强大的 PowerPoint 增强插件，提供多种实用工具来提升 PowerPoint 演示文稿的编辑效率和质量。

## 项目介绍

PPA（PowerPoint Advanced Add-in）是基于 .NET Framework 开发的 PowerPoint 插件，使用 NetOffice 库实现与 PowerPoint 的交互。该插件提供自定义功能区，包含多种实用工具，帮助用户更高效地创建和编辑演示文稿。

## 主要功能

### 核心功能

-   **自定义功能区集成** - 无缝集成到 PowerPoint 界面
-   **图形对齐和格式化工具** - 快速对齐、分布、吸附形状
-   **批量操作功能** - 批量格式化表格、文本、图表
-   **扩展格式化选项** - 可配置的格式化样式
-   **形状处理工具** - 形状创建、裁剪、显示/隐藏
-   **界面交互提示** - Toast 通知和进度指示

### 高级特性

-   **异步操作支持** - 耗时操作异步执行，不阻塞 UI
-   **配置化支持** - XML 配置文件自定义格式化样式
-   **多语言支持** - 支持中文（简体）和英文界面
-   **撤销优化** - 统一的撤销/重做管理
-   **快捷键支持** - 全局快捷键快速执行常用操作
-   **设置菜单** - 语言切换、参数配置、关于信息

## 技术栈

-   **C#** - 主要开发语言
-   **.NET Framework 4.8** - 目标框架
-   **NetOffice** - PowerPoint API 包装库
-   **Microsoft Office Interop** - 原生 Office COM 互操作
-   **Windows Forms** - UI 对话框支持
-   **Windows API** - 全局快捷键支持

## 项目结构

```
PPA/
├── AddIn/                    # 插件入口模块
│   ├── ThisAddIn.cs         # 插件主入口类
│   ├── ThisAddIn.Designer.cs
│   └── ThisAddIn.Designer.xml
│
├── Core/                     # 核心基础设施模块
│   ├── ExHandler.cs         # 异常处理类
│   ├── Profiler.cs          # 性能分析工具
│   └── ResourceManager.cs   # 多语言资源管理
│
├── Utilities/                # 通用工具模块
│   ├── Toast.cs             # 通知提示类
│   ├── FileLocator.cs       # 文件定位工具
│   └── ComListExtensions.cs # COM对象扩展方法
│
├── Formatting/               # 格式化业务模块
│   ├── FormatHelper.cs      # 格式化辅助工具
│   ├── BatchHelper.cs       # 批量操作辅助类
│   ├── FormattingConfig.cs  # 格式化配置管理
│   ├── AsyncOperationHelper.cs # 异步操作辅助类
│   ├── UndoHelper.cs        # 撤销操作管理
│   └── AlignHelper.cs       # 对齐辅助工具
│
├── Shape/                    # 形状处理模块
│   ├── ShapeUtils.cs        # 形状处理工具
│   └── MSOICrop.cs          # Microsoft Office交互裁剪类
│
├── UI/                       # UI模块
│   ├── CustomRibbon.cs      # 自定义功能区实现
│   ├── KeyboardShortcutHelper.cs # 全局快捷键管理
│   ├── Forms/               # 对话框窗体
│   │   ├── SettingsForm.cs # 设置对话框
│   │   └── AboutForm.cs    # 关于对话框
│   └── Ribbon.xml           # Ribbon配置文件
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
├── PPA.csproj               # 项目文件
├── PPA.sln                  # 解决方案文件
├── packages.config           # NuGet包配置
```

## 模块说明

### AddIn 模块

-   **ThisAddIn.cs**: 插件的主入口类，处理插件的初始化、资源管理和事件响应

### Core 模块

-   **ExHandler.cs**: 统一异常处理类，提供异常捕获、日志记录和性能监控功能
-   **Profiler.cs**: 性能监控类，提供方法执行时间测量、记录和日志功能
-   **ResourceManager.cs**: 多语言资源管理器，支持动态语言切换和本地化字符串管理

### Utilities 模块

-   **Toast.cs**: Toast 通知管理器，提供单消息框模式的用户提示
-   **FileLocator.cs**: 文件定位工具，在多个可能的位置搜索文件
-   **ComListExtensions.cs**: COM 对象列表扩展方法，提供批量释放功能

### Formatting 模块

-   **FormatHelper.cs**: 格式化辅助工具，提供表格、文本、图表的格式化功能
-   **BatchHelper.cs**: 批量操作辅助类，提供批量格式化、对齐等操作，支持异步执行
-   **FormattingConfig.cs**: 格式化配置管理，支持 XML 配置文件自定义格式化样式
-   **AsyncOperationHelper.cs**: 异步操作辅助类，提供统一的异步操作执行框架，支持进度报告
-   **UndoHelper.cs**: 撤销操作管理，统一管理 PowerPoint 撤销/重做操作
-   **AlignHelper.cs**: 对齐辅助工具，提供形状对齐、拉伸、吸附等相关操作

### Shape 模块

-   **ShapeUtils.cs**: 形状处理工具，提供形状创建、验证等实用方法
-   **MSOICrop.cs**: Microsoft Office 交互裁剪类，提供形状裁剪到幻灯片范围的功能

### UI 模块

-   **CustomRibbon.cs**: 自定义功能区实现，处理按钮点击和 UI 更新，支持动态本地化
-   **KeyboardShortcutHelper.cs**: 全局快捷键管理，使用 Windows API 实现系统级快捷键
-   **Forms/SettingsForm.cs**: 设置对话框，用于编辑格式化配置和切换语言
-   **Forms/AboutForm.cs**: 关于对话框，显示插件版本和项目信息
-   **Ribbon.xml**: Ribbon UI 配置文件，支持动态标签加载

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

-   **左对齐/右对齐/顶对齐/底对齐** - 快速对齐选中的形状
-   **水平居中/垂直居中** - 居中对齐形状
-   **横向分布/纵向分布** - 均匀分布多个形状
-   **吸附功能** - 快速吸附到参考线或页面边缘

#### 格式化工具

-   **美化表格** - 批量格式化表格样式（支持异步执行，显示进度）
-   **美化文本** - 批量格式化文本样式
-   **美化图表** - 批量格式化图表字体和样式
    -   快捷键：`Ctrl+3`（全局快捷键，可在任何窗口使用）

#### 形状工具

-   **插入形状** - 创建矩形外框或页面大小矩形
-   **隐显对象** - 隐藏选中对象或显示所有隐藏对象
-   **裁剪出框** - 将形状裁剪到幻灯片范围

#### 设置菜单

-   **语言切换** - 在中文（简体）和英文之间切换界面语言
-   **设置参数** - 编辑配置文件（`PPAConfig.xml`）
-   **关于** - 查看插件版本和项目信息

### 配置文件

插件会在 `%AppData%\PPA\` 目录（通常为 `C:\Users\<用户名>\AppData\Roaming\PPA\`）创建 `PPAConfig.xml` 配置文件，用于自定义格式化样式和快捷键设置。首次运行时会自动生成默认配置文件。

配置文件支持自定义：

-   表格样式（边框、填充、字体等）
-   文本样式（字体、大小、颜色等）
-   图表样式（标题、图例、坐标轴字体等）
-   快捷键设置（美化表格、美化文本、美化图表、插入形状等功能的快捷键）
-   格式：只需配置数字或字母（如 `"3"`, `"C"`, `"F1"`），系统会自动添加 `Ctrl` 修饰键
-   示例：`FormatChart="3"` 表示 `Ctrl+3`，`FormatTables="T"` 表示 `Ctrl+T`

### 快捷键

-   `Ctrl+3` - 美化图表（全局快捷键，可在 `PPAConfig.xml` 中自定义）
-   配置方式：在 `PPAConfig.xml` 的 `<Shortcuts>` 节点中，只需配置数字或字母（如 `FormatChart="3"`），系统会自动添加 `Ctrl` 修饰键
-   支持的键：数字 0-9、字母 A-Z、功能键 F1-F12

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

-   **Visual Studio 2019 或更高版本**
-   **.NET Framework 4.8**
-   **PowerPoint 2016 或更高版本**
-   **VSTO (Visual Studio Tools for Office)**

### 依赖项

-   NetOffice.PowerPointApi
-   Microsoft.Office.Interop.PowerPoint
-   Microsoft.Office.Tools.Common

## 项目结构说明

项目采用模块化设计，按功能划分为以下模块：

-   **Core** - 核心基础设施（异常处理、性能监控、资源管理）
-   **Utilities** - 通用工具（通知、文件定位、COM 对象扩展）
-   **Formatting** - 格式化业务逻辑（表格、文本、图表格式化）
-   **Shape** - 形状处理（形状创建、验证、裁剪）
-   **UI** - 用户界面（Ribbon、快捷键、对话框）
-   **AddIn** - 插件入口（初始化、生命周期管理）

## 开发指南

### 代码规范

-   遵循 C# 编码规范
-   使用 XML 文档注释
-   异常处理统一使用 `ExHandler.Run`
-   性能关键操作使用 `Profiler` 记录执行时间
-   用户提示使用 `Toast.Show`
-   所有用户可见文本使用 `ResourceManager.GetString` 进行本地化

### 添加新功能

1. 在相应的模块目录下创建新的类文件
2. 在 `CustomRibbon.cs` 中添加按钮处理逻辑
3. 在 `Ribbon.xml` 中添加 UI 定义
4. 在资源文件中添加本地化文本
5. 更新 `README.md` 文档

### 测试

-   在 Visual Studio 中按 F5 启动调试
-   PowerPoint 会自动启动并加载插件
-   检查日志输出以诊断问题

## 更新日志

### v1.0.0 (2025-01) - 第一版正式发布

**主要特性**：
-   ✅ 添加异步操作支持，提升大量表格格式化性能
-   ✅ 实现多语言支持（中文/英文）
-   ✅ 添加配置化支持，可通过 XML 自定义格式化样式
-   ✅ 优化撤销/重做功能
-   ✅ 添加全局快捷键支持（Ctrl+3 美化图表）
-   ✅ 添加设置菜单（语言切换、参数配置、关于）
-   ✅ 移除 VBA 依赖，提升性能和稳定性
-   ✅ 重构项目结构，采用模块化设计
-   ✅ 优化错误处理和日志记录
-   ✅ 移除配置文件版本号管理，简化配置结构
-   ✅ 移除旧路径配置文件迁移逻辑，统一使用新路径

## 贡献指南

欢迎提交 Issue 和 Pull Request 来帮助改进此项目。

### 贡献流程

1. Fork 本项目
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

### 代码提交规范

-   使用清晰的提交信息
-                确保代码通过编译和基本测试
-   更新相关文档

## 许可证

本项目采用 MIT 许可证。详见 LICENSE 文件。

## 项目链接

-   **GitHub**: [https://github.com/sandorn/PPA](https://github.com/sandorn/PPA)
-   **问题反馈**: [GitHub Issues](https://github.com/sandorn/PPA/issues)

## 致谢

感谢所有为这个项目做出贡献的开发者！

# PPA - PowerPoint增强插件

PPA是一个PowerPoint增强插件，提供多种实用功能来提升PowerPoint演示文稿的编辑效率和质量。

## 项目介绍

PPA（PowerPoint Advanced Add-in）是基于.NET开发的PowerPoint插件，使用NetOffice库实现与PowerPoint的交互。该插件提供自定义功能区，包含多种实用工具，帮助用户更高效地创建和编辑演示文稿。

## 主要功能

- 自定义功能区集成
- 图形对齐和格式化工具
- 批量操作功能
- 扩展格式化选项
- 形状处理工具
- 界面交互提示

## 技术栈

- C#
- .NET Framework
- NetOffice库
- Microsoft Office Interop

## 项目结构

```
PPA/
├── AlignHelper.cs           # 对齐辅助工具
├── BatchHelper.cs           # 批量操作辅助类
├── CustomRibbon.cs          # 自定义功能区实现
├── ExFormatter.cs           # 扩展格式化工具
├── ExHandler.cs             # 异常处理类
├── FormatHelper.cs          # 格式化辅助工具
├── MSOICrop.cs              # Microsoft Office交互裁剪类
├── PPA.csproj               # 项目文件
├── PPA.sln                  # 解决方案文件
├── Profiler.cs              # 性能分析工具
├── Properties/              # 属性文件夹
├── ShapeUtils.cs            # 形状处理工具
├── ThisAddIn.cs             # 插件主入口类
├── ToastN.cs                # 通知提示类
├── VBA.cs                   # VBA交互类
├── icon/                    # 图标资源文件夹
└── packages.config          # NuGet包配置
```

## 安装说明

1. 确保已安装PowerPoint
2. 构建项目生成DLL文件
3. 将插件DLL注册到PowerPoint
4. 重启PowerPoint后，插件将自动加载

## 使用方法

安装完成后，在PowerPoint界面中会出现自定义功能区。点击相应的按钮即可使用各项功能。

## 开发环境

- Visual Studio
- .NET Framework
- PowerPoint 开发环境

## 贡献指南

欢迎提交Issue和Pull Request来帮助改进此项目。

## 许可证

本项目采用MIT许可证。详见LICENSE文件。
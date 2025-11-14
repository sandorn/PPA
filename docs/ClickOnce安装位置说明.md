# ClickOnce VSTO 插件安装位置说明

## 📍 本地文件存储位置

使用 ClickOnce 部署的 VSTO 插件（如 PPA）安装后，文件存储在用户的**本地应用程序数据目录**中，而不是传统的 `Program Files` 目录。

### 默认安装路径

ClickOnce 应用程序通常安装在以下位置：

```
C:\Users\{用户名}\AppData\Local\Apps\2.0\{随机字符串}\{随机字符串}\{应用名}_{版本号}_{令牌}\
```

### 具体路径示例

对于 PPA 插件，完整路径可能类似：

```
C:\Users\YourName\AppData\Local\Apps\2.0\ABCD1234.EFG\HIJK5678.LMN\ppa...tion_0_9_0_0_abc123def456_0001.0000_zh-CN_abc123def456\
```

### 路径特点

1. **用户特定**：每个用户都有独立的安装目录
2. **版本隔离**：不同版本安装在不同的子目录中
3. **随机字符串**：路径中包含随机生成的字符串，用于安全隔离
4. **隐藏目录**：`AppData\Local` 是隐藏目录，需要在文件资源管理器中显示隐藏文件

## 🔍 如何查找插件安装位置

### 方法一：通过注册表查找

1. 按 `Win + R`，输入 `regedit`，打开注册表编辑器
2. 导航到：
   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PPA
   ```
3. 查看 `Manifest` 或 `LoadBehavior` 值，可能包含路径信息

### 方法二：通过 PowerPoint 信任中心

1. 打开 PowerPoint
2. 文件 → 选项 → 信任中心 → 信任中心设置
3. 受信任的加载项 → 查看已安装的加载项
4. 找到 PPA，查看其位置信息

### 方法三：通过 ClickOnce 应用程序清单

1. 打开文件资源管理器
2. 导航到：`C:\Users\{用户名}\AppData\Local\Apps\2.0\`
3. 搜索包含 "PPA" 或 "ppa" 的文件夹
4. 进入找到的文件夹，查找 `PPA.dll` 文件

### 方法四：使用 PowerShell 查找

在 PowerShell 中运行：

```powershell
# 查找所有包含 PPA 的 ClickOnce 应用程序
Get-ChildItem -Path "$env:LOCALAPPDATA\Apps\2.0" -Recurse -Filter "PPA.dll" -ErrorAction SilentlyContinue | Select-Object FullName

# 或者查找应用程序清单
Get-ChildItem -Path "$env:LOCALAPPDATA\Apps\2.0" -Recurse -Filter "*.application" -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*PPA*" } | Select-Object FullName
```

## 📂 目录结构说明

安装后的目录结构通常如下：

```
{安装根目录}/
├── PPA.dll                    # 主程序集
├── PPA.dll.manifest           # 程序集清单
├── PPA.dll.deploy              # ClickOnce 部署文件（如果存在）
├── PPA.application             # 应用程序清单
├── PPA.application.manifest    # 应用程序清单文件
├── en-US/                      # 英文资源文件夹
│   └── PPA.resources.dll.deploy
├── zh-CN/                      # 中文资源文件夹
│   └── PPA.resources.dll.deploy
├── UI/                         # UI 资源文件夹（如果使用文件系统）
│   └── Ribbon.xml.deploy
└── [其他依赖文件]
```

## ⚠️ 重要说明

### 1. 文件访问限制

- ClickOnce 应用程序的文件存储在用户目录中，**普通用户权限即可访问**
- 不需要管理员权限即可安装和卸载
- 文件被 ClickOnce 运行时管理，不建议手动修改

### 2. 文件命名规则

- 所有文件可能被重命名为 `.deploy` 扩展名（如 `PPA.dll.deploy`）
- 这是 ClickOnce 的安全机制，运行时会自动处理
- **注意**：PPA 插件已将 `Ribbon.xml` 改为嵌入式资源，不再需要文件系统访问

### 3. 多版本共存

- 不同版本的插件可以共存
- 每个版本安装在不同的目录中
- PowerPoint 会加载最新安装的版本

### 4. 卸载后的清理

- 卸载插件后，ClickOnce 会保留旧版本文件一段时间
- 可以通过"程序和功能"完全卸载
- 或者手动删除 `C:\Users\{用户名}\AppData\Local\Apps\2.0\` 下的相关文件夹

## 🔧 开发调试时的位置

### Debug 模式

在 Visual Studio 中按 F5 调试时，文件位置为：

```
D:\CODES\PPA\PPA\bin\Debug\
```

### Release 模式

直接运行 Release 版本时，文件位置为：

```
D:\CODES\PPA\PPA\bin\Release\
```

### ClickOnce 发布位置

发布文件位置（开发机器上）：

```
D:\CODES\PPA\PPA\bin\Release\publish\
```

## 📝 注册表信息

VSTO 插件在注册表中的位置：

```
HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PPA
```

注册表项包含：

- `Description` - 插件描述
- `FriendlyName` - 友好名称
- `LoadBehavior` - 加载行为（3 = 启动时加载）
- `Manifest` - 清单文件路径（指向 ClickOnce 安装位置）

## 🛠️ 常见问题

### Q: 为什么找不到插件 DLL 文件？

**A:** ClickOnce 应用程序的文件存储在隐藏目录中，需要：

1. 显示隐藏文件和文件夹
2. 使用上述方法之一查找

### Q: 可以手动修改插件文件吗？

**A:** 不建议手动修改，因为：

1. ClickOnce 会验证文件完整性
2. 修改后可能导致插件无法加载
3. 更新时会覆盖修改

### Q: 如何完全卸载插件？

**A:** 两种方法：

1. **通过控制面板**：
   - 控制面板 → 程序和功能
   - 找到 "PPA" 并卸载
2. **手动删除**：
   - 删除注册表项：`HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PPA`
   - 删除安装目录：`C:\Users\{用户名}\AppData\Local\Apps\2.0\{相关文件夹}`

### Q: 插件文件占用多少空间？

**A:** 通常很小，包括：

- 主 DLL：几百 KB 到几 MB
- 资源文件：几十到几百 KB
- 依赖项：如果未安装，可能需要额外空间

## 📚 相关资源

- [ClickOnce 部署概述](https://learn.microsoft.com/zh-cn/visualstudio/deployment/clickonce-security-and-deployment)
- [VSTO 插件部署](https://learn.microsoft.com/zh-cn/visualstudio/vsto/deploying-an-office-solution-by-using-clickonce)
- [ClickOnce 缓存位置](https://learn.microsoft.com/zh-cn/visualstudio/deployment/clickonce-cache-overview)

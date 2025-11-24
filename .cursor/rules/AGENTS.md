## PPA 代理通用规则

1. **输出风格**

   - 回复需简洁、聚焦结论，避免重复与无用寒暄。
   - 说明修复点时，优先描述“问题 → 方案 → 状态”。

2. **架构与结构**

   - **分层架构**：严格遵循以下目录职责：
     - `PPA/Core`: 核心基础设施（DI, Logging, ExHandler, Abstractions）。
     - `PPA/Utilities`: 通用工具类（ApplicationHelper, Toast, CommandExecutor）。
     - `PPA/UI`: 用户界面（Ribbon, Forms）。
     - `PPA/Shape`, `PPA/Manipulation`: 具体业务逻辑实现。
   - **异常处理**：核心业务逻辑必须使用 `ExHandler.Run` 包装，确保统一的异常捕获和日志记录。简单属性访问/设置可使用 `ExHandler.SafeGet` / `ExHandler.SafeSet`。

3. **依赖注入与服务访问**

   - **依赖注入**：服务注册优先在 `PPA.Core.DI.ServiceCollectionExtensions` 中进行。
   - **构造注入**：新类优先通过构造函数声明依赖（如 `IShapeHelper`, `ICommandExecutor`）。
   - **静态/遗留代码**：在无法使用构造注入的场景（如静态 Helper），使用 `ApplicationProvider.Current.ServiceProvider` 获取服务，或使用 `LoggerProvider.GetLogger()` 获取日志。

4. **COM/NetOffice 规范**

   - **Application 获取**：业务方法应通过参数接收 `NETOP.Application` 实例，并使用 `ApplicationHelper.EnsureValidNetApplication` 进行校验。禁止直接访问 `Globals.ThisAddIn`。
   - **资源释放**：
     - 显式使用 `using` 块管理 `NETOP.Shape`, `NETOP.Selection`, `NETOP.ShapeRange` 等 COM 对象。
     - 在批量操作中，遵循“收集 → 处理 → 释放”模式，确保循环中创建的临时 COM 对象及时释放（参考 `MSOICrop.cs`）。
   - **版本兼容**：使用 `Application.Version` 检查功能兼容性（如布尔运算对 2013+ 的依赖）。

5. **日志与用户反馈**

   - **日志记录**：统一使用 `ILogger` 接口。异常详情由 `ExHandler` 自动记录，业务流程中的关键节点需手动记录 Info/Warning。
   - **用户反馈**：使用 `Toast.Show` 显示操作结果（成功/警告/错误），支持多语言（`ResourceManager.GetString`）。禁止使用 `MessageBox` 阻断用户操作。

6. **业务逻辑模式**

   - **选区验证**：操作前必须通过 `IShapeHelper.ValidateSelection` (或类似 Helper) 验证选区有效性。
   - **批量操作**：对于多对象操作，采用“快照状态 → 执行变更 → 结果反馈”流程，必要时保存 Z-Order 或其他状态以备恢复。

7. **文档与配置**

   - 所有架构调整或新服务引入需同步更新 `docs/` 下的相关文档。
   - 新增配置项需在 `PPAConfig.xml` 相关说明中记录默认值和用途。

8. **日期处理**
   - 在处理任何与日期和时间相关的任务时，你必须使用占位符，如 `[当前日期]`。
   - 你不能自己编造日期。

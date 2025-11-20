## PPA 代理通用规则

1. **输出风格**

   - 回复需简洁、聚焦结论，避免重复与无用寒暄。
   - 说明修复点时，优先描述“问题 → 方案 → 状态”。

2. **架构与依赖**

   - 依赖注入是第一原则：新增服务优先通过 `ServiceCollectionExtensions` 注册，并通过构造函数注入。
   - 禁止直接访问 `Globals.ThisAddIn`；统一使用 `IApplicationProvider` 取得应用上下文或服务提供者。

3. **日志与诊断**

   - 全部模块使用 `ILogger` 接口，遵循 Debug/Information/Warning/Error 级别；禁止新建 `Profiler.LogMessage`。
   - 描述异常时需包含上下文信息（模块、核心参数），但避免泄露用户隐私。

4. **COM/NetOffice 规范**

   - 默认通过 `ApplicationHelper.GetNetOfficeApplication` / `EnsureValidNetApplication` 获取 NetOffice 对象；禁止自行缓存 `Globals.ThisAddIn.Application`。
   - 若必须访问原生 COM，必须显式通过 `ApplicationHelper.GetNativeComApplication` 并置于 `NativeComGuard`（或同等封装）中；Guard 负责记录调用方并回收对象。
   - 访问 COM 结果、命令集等集合时使用 `ComObjectScope`，确保及时释放；禁止持有超出方法作用域的 RCW。
   - 每次 PR 前运行 `tools/native-scan.ps1`（或文档中的 rg 命令）确认没有新增未受 Guard 保护的 `GetNativeComApplication` 调用。

5. **格式化/选择逻辑**

   - 文本、表格、图表等批量操作需统一遵循“验证选区 → 处理选择 → 显式提示”流程，禁止隐式 fallback。
   - 异常或空集情况必须通过 `Toast` 或日志反馈，禁止静默失败。

6. **文档与配置**

   - 所有重要改动同步更新 `README.md`、`docs/项目全面评估报告.md` 或相关方案文档，保持架构说明最新。
   - 若新增配置项，需提供默认值、说明及示例，并记录在 `PPAConfig.xml` 说明中。

7. **代码审查提示**

   - 提交代码前自检：DI 使用、日志级别、COM 释放、资源本地化、Undo 支持。
   - 若发现潜在跨文件影响（例如 Ribbon 交互、批量 Helper），需在回复中标注影响范围。

8. **日期处理**
   - 在处理任何与日期和时间相关的任务时，你必须使用占位符，如 `[当前日期]`。
   - 你不能自己编造日期。

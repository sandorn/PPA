using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace Project.Utilities
{
    /// <summary>
    /// 统一异常处理类
    /// 提供异常捕获、日志记录和性能监控功能
    /// </summary>
    public static class ExHandler
    {
        #region Properties

        /// <summary>
        /// 是否启用操作耗时记录
        /// 默认为false以提升性能
        /// </summary>
        public static bool EnableTiming { get; set; } = false;

        /// <summary>
        /// 日志文件路径
        /// 统一使用Profiler的日志路径配置，确保日志一致性
        /// </summary>
        public static string LogFilePath
        {
            get { return Profiler.LogFilePath; }
            set { Profiler.LogFilePath = value; }
        }

        /// <summary>
        /// 初始化 ExHandler 的全局配置
        /// </summary>
        /// <param name="enableFileLogging">是否启用文件日志</param>
        /// <param name="enableTiming">是否启用性能监控</param>
        /// <param name="logFilePath">日志文件路径</param>
        public static void Initialize(bool enableFileLogging = false, bool enableTiming = false)
        {
            EnableTiming = enableTiming;
        }

        #endregion Properties

        #region Methods

        // 无返回值方法（带调用方法名捕获）
        public static void Run(
            Action action,
            string context = null,
            bool? enableTiming = null, // 局部覆盖参数
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "")
        {
            TimeSpan elapsed = TimeSpan.Zero;
            // 决策逻辑：优先使用局部参数，否则使用全局默认值
            bool shouldTime = enableTiming ?? EnableTiming;

            try
            {
                if (shouldTime)
                {
                    elapsed = Profiler.Time(action, callerMethod, callerFile);
                }
                else
                {
                    action();
                    Profiler.LogMessage(context, "ExHandler", callerMethod, callerFile);
                }
            }
            catch (Exception ex)
            {
                HandleException(ex,
                    effectiveContext: context ?? callerMethod,
                    callerMethod: callerMethod,
                    callerFile: callerFile);
            }
        }

        // 有返回值方法（带调用方法名捕获）
        public static T Run<T>(
            Func<T> func,
            string context = null,
            bool? enableTiming = null, // 局部覆盖参数
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string callerFile = "",
            T defaultValue = default)
        {
            TimeSpan elapsed = TimeSpan.Zero;
            T result = defaultValue;
            // 决策逻辑：优先使用局部参数，否则使用全局默认值
            bool shouldTime = enableTiming ?? EnableTiming;

            try
            {
                if (shouldTime)
                {
                    (result, elapsed) = Profiler.Time(func, callerMethod, callerFile);
                }
                else
                {
                    result = func();
                    Profiler.LogMessage(context, "ExHandler", callerMethod, callerFile);
                }

                return result;
            }
            catch (Exception ex)
            {
                HandleException(ex,
                    effectiveContext: context ?? callerMethod,
                    callerMethod: callerMethod,
                    callerFile: callerFile);
                return defaultValue;
            }
        }

        // 获取实际抛出异常的方法名
        private static string GetActualMethodName(Exception ex)
        {
            try
            {
                // 从堆栈中获取第一个非系统方法
                var stackTrace = new StackTrace(ex, fNeedFileInfo: true);
                foreach (StackFrame frame in stackTrace.GetFrames())
                {
                    var method = frame.GetMethod();
                    if (method == null) continue;

                    // 跳过系统方法
                    var declaringType = method.DeclaringType;
                    if (declaringType == null) continue;

                    if (declaringType.Namespace?.StartsWith("System.") != false ||
                        declaringType.Namespace.StartsWith("Microsoft."))
                    {
                        continue;
                    }
                    return $"{declaringType.Name}.{method.Name}";
                }
            }
            catch { /* 安全捕获 */ }

            return null;
        }

        /// <summary>
        /// 统一异常处理方法
        /// 记录异常信息、调用位置、耗时等详细信息
        /// </summary>
        /// <param name="ex">捕获的异常</param>
        /// <param name="effectiveContext">操作上下文</param>
        /// <param name="callerMethod">调用方法名</param>
        /// <param name="callerFile">调用文件路径</param>
        private static void HandleException(Exception ex, string effectiveContext, string callerMethod, string callerFile)
        {
            // 获取调用者类名
            var callerClass = Path.GetFileNameWithoutExtension(callerFile);

            // 获取当前方法名（实际抛出异常的方法）
            var actualMethod = GetActualMethodName(ex) ?? "未知方法";

            try
            {
                // 输出调试信息
                Profiler.LogMessage(ExFormatter.FormatFullException(ex), "ExHandler", callerMethod, callerFile);
            }
            catch {/* 防止日志失败导致二次异常 */}
        }

        #endregion Methods
    }

    public static class ExFormatter
    {
        #region Methods

        public static string FormatFullException(Exception ex)
        {
            if (ex == null) return string.Empty;

            var sb = new StringBuilder();
            AppendExceptionDetails(sb, ex, depth: 0);
            return sb.ToString();
        }

        private static void AppendExceptionDetails(StringBuilder sb, Exception ex, int depth)
        {
            if (depth > 0) sb.Append('\n').Append(' ', depth * 2);

            sb.Append($"[{ex.GetType().Name}] {ex.Message}");
            sb.Append($"\n{"HResult:",-10} 0x{ex.HResult:X8}");

            if (!string.IsNullOrWhiteSpace(ex.StackTrace))
            {
                sb.Append($"\n{"Stack Trace:",-10}");
                sb.Append(FormatStackTrace(ex.StackTrace));
            }

            if (ex.InnerException != null)
            {
                sb.Append($"\n{"Inner:",-10}");
                AppendExceptionDetails(sb, ex.InnerException, depth + 1);
            }
        }

        private static string FormatStackTrace(string stackTrace)
        {
            var lines = stackTrace.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
            return "\n          " + string.Join("\n          ", lines);
        }

        #endregion Methods
    }

}

/*
// ExHandler 使用示例

// 1. 基本使用 - 自动捕获异常和调用信息
public void BasicUsage()
{
    // 无返回值方法
    ExHandler.Run(() => {
        // 可能抛出异常的代码
        ProcessData();
    }, "数据处理操作");
    
    // 有返回值方法
    var result = ExHandler.Run(() => {
        // 可能抛出异常的代码
        return CalculateResult();
    }, "计算结果");
}

// 2. 配置全局设置
public void ConfigureGlobalSettings()
{
    // 在应用程序启动时配置
    ExHandler.EnableFileLogging = true;
    ExHandler.EnableTiming = true; // 记录所有操作的耗时
    ExHandler.LogFilePath = "app_errors.log"; // 设置日志文件路径
}

// 3. 局部覆盖全局设置
public void LocalOverride()
{
    // 全局设置为不记录耗时
    ExHandler.EnableTiming = false;
    
    // 但对特定操作启用计时
    ExHandler.Run(() => {
        // 这个操作会被计时
        PerformCriticalOperation();
    }, "关键操作", enableTiming: true);
    
    // 这个操作不会被计时
    ExHandler.Run(() => {
        // 这个操作不会被计时
        PerformNormalOperation();
    }, "普通操作");
}

// 4. 处理不同类型的异常
public void HandleDifferentExceptions()
{
    // 处理可能抛出的不同类型异常
    var result = ExHandler.Run(() => {
        if (someCondition) {
            throw new InvalidOperationException("无效操作");
        }
        
        if (anotherCondition) {
            throw new ArgumentException("参数错误");
        }
        
        return ProcessData();
    }, "数据处理");
}

// 5. 嵌套调用
public void NestedCalls()
{
    ExHandler.Run(() => {
        // 外层操作
        ExHandler.Run(() => {
            // 内层操作
            DoInnerWork();
        }, "内层操作");
        
        DoOuterWork();
    }, "外层操作");
}

// 6. 与Profiler结合使用
public void WithProfiler()
{
    // ExHandler内部会调用Profiler进行计时
    ExHandler.Run(() => {
        // 也可以单独使用Profiler进行更细粒度的计时
        Profiler.Time(() => {
            // 特定部分的性能监控
            PerformCriticalSection();
        });
        
        DoOtherWork();
    }, "完整操作");
}

// 7. 自定义默认返回值
public void CustomDefaultValues()
{
    // 为不同类型指定默认返回值
    var stringValue = ExHandler.Run(() => {
        if (errorCondition) {
            throw new Exception("处理失败");
        }
        return "成功结果";
    }, "字符串处理", defaultValue: "默认值");
    
    var intValue = ExHandler.Run(() => {
        if (errorCondition) {
            throw new Exception("计算失败");
        }
        return 42;
    }, "数值计算", defaultValue: -1);
}

// 8. 在异步方法中使用
public async Task AsyncUsage()
{
    // 注意：当前ExHandler不支持异步，需要使用Task.Run包装
    var result = await ExHandler.Run(() => Task.Run(async () => {
        // 异步操作
        return await ProcessDataAsync();
    }, "异步数据处理")).Result;
    
    // 或者使用更直接的方式
    var result2 = await Task.Run(() => ExHandler.Run(() => {
        // 同步包装异步操作
        return ProcessDataAsync().GetAwaiter().GetResult();
    }, "同步包装异步操作"));
}

// 9. 批量处理
public void BatchProcessing()
{
    var items = GetItemsToProcess();
    var results = new List<Result>();
    
    foreach (var item in items)
    {
        // 每个项目的处理都被单独捕获异常
        var result = ExHandler.Run(() => {
            return ProcessItem(item);
        }, $"处理项目 {item.Id}", defaultValue: new Result());
        
        if (result != null)
        {
            results.Add(result);
        }
    }
    
    Profiler.LogMessage($"成功处理 {results.Count}/{items.Count} 个项目");
}

// 10. 性能监控和异常分析
public void PerformanceAndExceptionAnalysis()
{
    // 启用计时和文件日志
    ExHandler.EnableTiming = true;
    ExHandler.EnableFileLogging = true;
    
    // 执行操作
    ExHandler.Run(() => {
        // 复杂业务逻辑
        ProcessComplexBusinessLogic();
    }, "复杂业务逻辑处理");
    
    // 检查日志文件以分析性能和异常
    Profiler.LogMessage("请查看日志文件以分析性能和异常信息");
}
 */
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
		/// 是否启用文件日志记录
		/// </summary>
		public static bool EnableFileLogging { get; set; } = false;

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
		public static void Initialize(bool enableFileLogging = false,bool enableTiming = false)
		{
			EnableFileLogging=enableFileLogging;
			EnableTiming=enableTiming;
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
				if(shouldTime)
				{
					elapsed = Profiler.Time(action,$"{Path.GetFileName(callerFile)} | {callerMethod}");
				} else
				{
					action();
					Debug.WriteLine($"[ExHandler] {context ?? callerMethod} 调用文件: {Path.GetFileName(callerFile)}");
				}
			} catch(Exception ex)
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
				if(shouldTime)
				{
					(result, elapsed) = Profiler.Time(func,$"{Path.GetFileName(callerFile)} | {callerMethod}");
				} else
				{
					result = func();
					Debug.WriteLine($"[ExHandler] {context ?? callerMethod} 调用文件: {Path.GetFileName(callerFile)}");
				}

				return result;
			} catch(Exception ex)
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
				var stackTrace = new StackTrace(ex,fNeedFileInfo: true);
				foreach(StackFrame frame in stackTrace.GetFrames())
				{
					var method = frame.GetMethod();
					if(method == null) continue;

					// 跳过系统方法
					var declaringType = method.DeclaringType;
					if(declaringType == null) continue;

					if(declaringType.Namespace?.StartsWith("System.") != false ||
						declaringType.Namespace.StartsWith("Microsoft."))
					{
						continue;
					}
					return $"{declaringType.Name}.{method.Name}";
				}
			} catch { /* 安全捕获 */ }

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
		private static void HandleException(Exception ex,string effectiveContext,string callerMethod,string callerFile)
		{
			// 获取调用者类名
			var callerClass = Path.GetFileNameWithoutExtension(callerFile);

			// 获取当前方法名（实际抛出异常的方法）
			var actualMethod = GetActualMethodName(ex) ?? "未知方法";

			// 构建日志内容
			string logContent = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}]\t[{effectiveContext}] 出错！";


			logContent += $"\n调用位置: {callerClass}.{callerMethod}" + $"\n异常位置: {actualMethod}" + $"\n{ExFormatter.FormatFullException(ex)}";

			try
			{
				// 输出调试信息
				Debug.WriteLine("##########\n### 操作失败 ###");
				Debug.WriteLine(logContent + "\n##########\n");

				// 文件日志
				if(EnableFileLogging)
				{
					File.AppendAllText(LogFilePath,logContent + "##########\n");
				}
			} catch {/* 防止日志失败导致二次异常 */}
		}

		#endregion Methods
	}

	public static class ExFormatter
	{
		#region Methods

		public static string FormatFullException(Exception ex)
		{
			if(ex==null) return string.Empty;

			var sb = new StringBuilder();
			AppendExceptionDetails(sb,ex,depth: 0);
			return sb.ToString();
		}

		private static void AppendExceptionDetails(StringBuilder sb,Exception ex,int depth)
		{
			if(depth>0) sb.Append('\n').Append(' ',depth*2);

			sb.Append($"[{ex.GetType().Name}] {ex.Message}");
			sb.Append($"\n{"HResult:",-10} 0x{ex.HResult:X8}");

			if(!string.IsNullOrWhiteSpace(ex.StackTrace))
			{
				sb.Append($"\n{"Stack Trace:",-10}");
				sb.Append(FormatStackTrace(ex.StackTrace));
			}

			if(ex.InnerException!=null)
			{
				sb.Append($"\n{"Inner:",-10}");
				AppendExceptionDetails(sb,ex.InnerException,depth+1);
			}
		}

		private static string FormatStackTrace(string stackTrace)
		{
			var lines = stackTrace.Split(['\r','\n'],StringSplitOptions.RemoveEmptyEntries);
			return "\n          "+string.Join("\n          ",lines);
		}

		#endregion Methods
	}

}

/*
// 在应用程序初始化时配置
ExHandler.EnableFileLogging = true;
ExHandler.EnableTiming = true; // 记录所有操作的耗时
ExHandler.LogFilePath = true; // 在异常中记录耗时

// 使用示例
ExceptionHandler.Run(() =>
{
    // 可能抛出异常的代码
    ProcessData();
}, "数据处理操作");

// 带返回值的使用
var result = ExceptionHandler.Run(() =>
{
    return CalculateResult();
}, "计算结果");
 */
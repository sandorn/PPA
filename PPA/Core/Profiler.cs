using PPA.Core.Abstraction.Infrastructure;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace PPA.Core
{
	/// <summary>
	/// 性能监控类 提供方法执行时间测量、记录和日志功能
	/// </summary>
	public static class Profiler
	{
		#region Public Properties

		/// <summary>
		/// 是否启用文件日志记录 默认为false以避免不必要的文件IO操作
		/// </summary>
		public static bool EnableFileLogging { get; set; } = false;

		/// <summary>
		/// 性能日志文件路径
		/// </summary>
		public static string LogFilePath { get; set; } = "Profiler.log";

		/// <summary>
		/// 最多保留的日志文件数量
		/// </summary>
		public static int MaxLogFiles { get; set; } = 10;

		/// <summary>
		/// 日志文件最长保留时间
		/// </summary>
		public static TimeSpan? MaxLogAge { get; set; } = TimeSpan.FromDays(7);

		/// <summary>
		/// 最小写入日志级别（低于该级别的日志会被过滤）
		/// </summary>
		public static LogLevel MinimumLogLevel { get; set; } = LogLevel.Information;

		#endregion Public Properties

		#region Private Fields

		private const int BufferCapacity = 20; // 日志缓冲区容量
		private static readonly Queue<string> _buffer = new(); // 日志缓冲区
		private static readonly object _lockObj = new(); // 线程同步锁
		private static StreamWriter _writer; // 文件写入器
		private static readonly TimeSpan MaxFlushInterval = TimeSpan.FromSeconds(1);
		private static DateTime _lastFlushTime = DateTime.UtcNow;

		#endregion Private Fields

		#region Public Methods

		/// <summary>
		/// 测量无返回值方法的执行时间 自动记录性能数据到调试输出和可选的文件日志
		/// </summary>
		/// <param name="action"> 要执行的操作 </param>
		/// <param name="message"> 消息内容 </param>
		/// <param name="logLevel"> 日志级别（如：INFO, WARN, ERROR等） </param>
		/// <param name="methodIdentifier"> 方法标识符（如果提供则使用，否则自动获取） </param>
		/// <param name="callerMethod"> 方法名称（默认为调用者方法名） </param>
		/// <param name="callerFile"> 调用者文件路径（默认为调用者文件路径） </param>
		/// <returns> 执行耗时 </returns>
		public static TimeSpan Time(Action action,string message = null,string logLevel = "INFO",string methodIdentifier = null,[CallerMemberName] string callerMethod = "",[CallerFilePath] string callerFile = "")
		{
			var sw = Stopwatch.StartNew();
			action();
			sw.Stop();

			// 直接调用LogMessage方法记录性能数据
			message=$"{message} 执行耗时: {sw.Elapsed.TotalMilliseconds:F3} ms";
			if(string.IsNullOrEmpty(methodIdentifier))
			{
				methodIdentifier=string.IsNullOrEmpty(callerFile)
					? callerMethod
					: $"{Path.GetFileNameWithoutExtension(callerFile)}.{callerMethod}";
			}
			LogMessage(message,logLevel,methodIdentifier,callerMethod,callerFile);
			return sw.Elapsed;
		}

		/// <summary>
		/// 测量有返回值方法的执行时间 自动记录性能数据到调试输出和可选的文件日志
		/// </summary>
		/// <typeparam name="T"> 返回值类型 </typeparam>
		/// <param name="func"> 要执行的函数 </param>
		/// <param name="message"> 消息内容 </param>
		/// <param name="logLevel"> 日志级别（如：INFO, WARN, ERROR等） </param>
		/// <param name="methodIdentifier"> 方法标识符（如果提供则使用，否则自动获取） </param>
		/// <param name="callerMethod"> 方法名称（默认为调用者方法名） </param>
		/// <param name="callerFile"> 调用者文件路径（默认为调用者文件路径） </param>
		/// <returns> 元组：方法返回值和执行耗时 </returns>
		public static (T result, TimeSpan elapsed) Time<T>(Func<T> func,string message = null,string logLevel = "INFO",string methodIdentifier = null,[CallerMemberName] string callerMethod = "",[CallerFilePath] string callerFile = "")
		{
			var sw = Stopwatch.StartNew();
			var result = func();
			sw.Stop();

			// 直接调用LogMessage方法记录性能数据
			message=$"{message} 执行耗时: {sw.Elapsed.TotalMilliseconds:F3} ms";
			if(string.IsNullOrEmpty(methodIdentifier))
			{
				methodIdentifier=string.IsNullOrEmpty(callerFile)
					? callerMethod
					: $"{Path.GetFileNameWithoutExtension(callerFile)}.{callerMethod}";
			}
			LogMessage(message,logLevel,methodIdentifier,callerMethod,callerFile);
			return (result, sw.Elapsed);
		}

		#endregion Public Methods

		/// <summary>
		/// 根据当前策略清理多余或过期的日志文件。
		/// </summary>
		/// <param name="directory"> 日志目录。 </param>
		/// <param name="searchPattern"> 匹配的日志文件模式。 </param>
		public static void CleanupLogFiles(string directory,string searchPattern = "PPA_*.log")
		{
			try
			{
				if(string.IsNullOrEmpty(directory)||!Directory.Exists(directory))
				{
					return;
				}

				var files = new DirectoryInfo(directory)
					.GetFiles(searchPattern,SearchOption.TopDirectoryOnly)
					.OrderByDescending(f => f.CreationTimeUtc)
					.ToList();

				if(files.Count==0)
				{
					return;
				}

				var maxFiles = MaxLogFiles>0 ? MaxLogFiles : int.MaxValue;
				var maxAge = MaxLogAge;

				for(int i = 0;i<files.Count;i++)
				{
					var file = files[i];
					bool exceedsCount = i>=maxFiles;
					bool exceedsAge = maxAge.HasValue && (DateTime.UtcNow-file.CreationTimeUtc)>maxAge.Value;

					if(exceedsCount||exceedsAge)
					{
						try
						{
							file.Delete();
						} catch
						{
							// 忽略删除失败
						}
					}
				}
			} catch
			{
				// 忽略清理异常
			}
		}

		/// <summary>
		/// 强制把缓冲日志写入文件
		/// </summary>
		public static void Flush()
		{
			if(!EnableFileLogging)
			{
				return;
			}

			lock(_lockObj)
			{
				FlushBuffer();
			}
		}

		#region Public Methods

		/// <summary>
		/// 记录自定义日志信息 开发状态下输出到Debug控制台，非开发状态下写入日志文件
		/// </summary>
		/// <param name="message"> 日志消息内容 </param>
		/// <param name="logLevel"> 日志级别（如：INFO, WARN, ERROR等） </param>
		/// <param name="methodIdentifier"> 可选的调用位置标识符（如果提供则使用，否则自动获取） </param>
		/// <param name="callerMethod"> 调用者方法名（自动获取，当 methodIdentifier 为空时使用） </param>
		/// <param name="callerFile"> 调用者文件路径（自动获取，当 methodIdentifier 为空时使用） </param>
		public static void LogMessage(string message,string logLevel = "INFO",string methodIdentifier = null,[CallerMemberName] string callerMethod = "",[CallerFilePath] string callerFile = "")
		{
			// 日志级别过滤
			var level = ParseLogLevelString(logLevel);
			if(level<MinimumLogLevel)
			{
				return;
			}

			// 如果未提供 methodIdentifier，则自动构建
			if(string.IsNullOrEmpty(methodIdentifier))
			{
				methodIdentifier=string.IsNullOrEmpty(callerFile)
					? callerMethod
					: $"{Path.GetFileNameWithoutExtension(callerFile)}.{callerMethod}";
			}

			var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}]\t[{logLevel}]\t{methodIdentifier}\t{message}";

#if DEBUG
			// 开发状态：输出到Debug控制台
			Debug.WriteLine($"[{logLevel}]\t{methodIdentifier}\t{message}");
#endif

			// 如果启用了文件日志，也写入文件（DEBUG 和 Release 都支持）
			if(EnableFileLogging)
			{
				lock(_lockObj)
				{
					_buffer.Enqueue(line);

					var shouldFlush = _buffer.Count >= BufferCapacity || _writer == null || (DateTime.UtcNow - _lastFlushTime) >= MaxFlushInterval;
					if(shouldFlush)
					{
						FlushBuffer();
					}
				}
			}
		}

		#endregion Public Methods

		#region Private Methods

		private static void FlushBuffer()
		{
			if(_writer==null)
			{
				try
				{
					// 延迟初始化写入器
					_writer=new StreamWriter(LogFilePath,append: true)
					{
						AutoFlush=true
					};
				} catch
				{
					// 初始化失败时清空缓冲区
					_buffer.Clear();
					return;
				}
			}

			try
			{
				// 写入所有缓冲日志
				while(_buffer.Count>0)
				{
					_writer.WriteLine(_buffer.Dequeue());
				}
				_writer.Flush();
				_lastFlushTime=DateTime.UtcNow;
			} catch
			{
				// 写入失败时清理资源
				_writer.Dispose();
				_writer=null;
				_buffer.Clear();
			}
		}

		private static LogLevel ParseLogLevelString(string logLevel)
		{
			if(string.IsNullOrWhiteSpace(logLevel))
			{
				return LogLevel.Information;
			}

			var text = logLevel.Trim();

			if(int.TryParse(text,out var numeric))
			{
				if(numeric>=(int) LogLevel.Debug&&numeric<=(int) LogLevel.Error)
				{
					return (LogLevel) numeric;
				}
			}

			switch(text.ToUpperInvariant())
			{
				case "DEBUG":
					return LogLevel.Debug;

				case "INFO":
				case "INFORMATION":
					return LogLevel.Information;

				case "WARN":
				case "WARNING":
					return LogLevel.Warning;

				case "ERROR":
					return LogLevel.Error;
			}

			if(Enum.TryParse<LogLevel>(text,true,out var parsed))
			{
				return parsed;
			}

			return LogLevel.Information;
		}

		#endregion Private Methods
	}
}

namespace PPA.Core.Extensions
{
	// 在Profiler类中添加扩展方法
	public static class ProfilerEx
	{
		#region Public Methods

		/// <summary>
		/// [扩展方法] 测量操作执行时间 - 使用流畅API风格
		/// </summary>
		/// <remarks> 注意：此方法会使所有Action获得Time()方法 </remarks>
		public static TimeSpan Time(
			this Action action,
			string logLevel = "INFO",
			string methodIdentifier = null,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "")
		{
			return Profiler.Time(action,logLevel,methodIdentifier,callerMethod,callerFile);
		}

		/// <summary>
		/// [扩展方法] 测量操作执行时间 - 使用流畅API风格
		/// </summary>
		/// <remarks> 注意：此方法会使所有Func获得Time()方法 </remarks>
		public static (TResult Result, TimeSpan Elapsed) Time<TResult>(
			this Func<TResult> func,
			string logLevel = "INFO",
			string methodIdentifier = null,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "")
		{
			return Profiler.Time(func,logLevel,methodIdentifier,callerMethod,callerFile);
		}

		// 可选：添加常用参数类型的重载
		public static TimeSpan Time<Targs>(
			this Action<Targs> action,
			Targs arg,
			string logLevel = "INFO",
			string methodIdentifier = null,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "")
		{
			return Time(() => action(arg),logLevel,methodIdentifier,callerMethod,callerFile);
		}

		public static (TResult Result, TimeSpan Elapsed) Time<Targs, TResult>(
			this Func<Targs,TResult> func,
			Targs arg,
			string logLevel = "INFO",
			string methodIdentifier = null,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "")
		{
			return Time(() => func(arg),logLevel,methodIdentifier,callerMethod,callerFile);
		}

		/// <summary>
		/// [扩展方法] 记录自定义日志信息 开发状态下输出到Debug控制台，非开发状态下写入日志文件
		/// </summary>
		/// <param name="_"> 任意对象（仅用于扩展方法语法，实际未使用） </param>
		/// <param name="message"> 日志消息内容 </param>
		/// <param name="logLevel"> 日志级别（如：INFO, WARN, ERROR等） </param>
		/// <param name="methodIdentifier"> 方法标识符（如果提供则使用，否则自动获取） </param>
		/// <param name="callerMethod"> 调用者方法名（自动获取） </param>
		/// <param name="callerFile"> 调用者文件路径（自动获取） </param>
		public static void Log(this object _,string message,string logLevel = "INFO",string methodIdentifier = null,[CallerMemberName] string callerMethod = "",[CallerFilePath] string callerFile = "")
		{
			if(string.IsNullOrEmpty(methodIdentifier))
			{
				methodIdentifier=string.IsNullOrEmpty(callerFile)
					? callerMethod
					: $"{Path.GetFileNameWithoutExtension(callerFile)}.{callerMethod}";
			}
			Profiler.LogMessage(message,logLevel,methodIdentifier);
		}

		#endregion Public Methods
	}
}

/*
// ============================================= 超级简单使用示例（放在代码文件末尾即可） =============================================

// 示例1：基本用法
Profiler.Time(() =>
{
    // 你的代码放在这里
    Thread.Sleep(100);
});
Profiler.Time(() => { ...... });

// 示例2：带返回值的方法
var (result, time) = Profiler.Time(() =>
{
    return "计算结果";
});

var (result, time) = Profiler.Time(() => 42);

// 示例3：ProfilerEx 使用扩展方法（更简洁）
Action myAction = () => { ...... };
myAction.Time();

Func<string> myFunc = () => "hello";
var (data, elapsed) = myFunc.Time();

// 示例4：实际使用场景 在方法开始时测量性能
public void MyMethod()
{
    Profiler.Time(() =>
    {
        // 方法的具体实现
        DoWork();
        ProcessData();
    });
}

// 启用文件日志：
Profiler.EnableFileLogging = true;

// 直接调用静态方法
Profiler.LogMessage("这是一条信息日志");
Profiler.LogMessage("这是一条错误日志", "ERROR");

// 使用扩展方法
var myObject = new SomeClass();
myObject.Log("这是通过扩展方法记录的日志");
myObject.Log("这是一条警告日志", "WARN");

 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using ToastAPI;

namespace Project.Utilities
{
	/// <summary>
	/// 性能监控类
	/// 提供方法执行时间测量、记录和日志功能
	/// </summary>
	public static class Profiler
	{
		#region Public Properties

		/// <summary>
		/// 是否启用文件日志记录
		/// 默认为false以避免不必要的文件IO操作
		/// </summary>
		public static bool EnableFileLogging { get; set; } = false;

		/// <summary>
		/// 性能日志文件路径
		/// </summary>
		public static string LogFilePath { get; set; } = "Profiler.log";

		#endregion Public Properties

		#region Private Fields

		private const int BufferCapacity = 100; // 日志缓冲区容量
		private static readonly Queue<string> _buffer = new(); // 日志缓冲区
		private static readonly object _lockObj = new(); // 线程同步锁
		private static StreamWriter _writer; // 文件写入器

		#endregion Private Fields

		#region Public Methods

		/// <summary>
		/// 测量无返回值方法的执行时间
		/// 自动记录性能数据到调试输出和可选的文件日志
		/// </summary>
		/// <param name="action">要执行的操作</param>
		/// <param name="callerMethod">方法名称（默认为调用者方法名）</param>
		/// <param name="filePath">调用者文件路径（默认为调用者文件路径）</param>
		/// <returns>执行耗时</returns>
		public static TimeSpan Time(Action action,[CallerMemberName] string callerMethod = "", [CallerFilePath] string filePath = "")
		{
			var sw = Stopwatch.StartNew();
			action();
			sw.Stop();

			// 构建方法标识符，包含文件名和方法名
			string methodIdentifier = string.IsNullOrEmpty(filePath) 
				? callerMethod 
				: $"{Path.GetFileName(filePath)} | {callerMethod}";
			
			LogPerformance(methodIdentifier, sw.Elapsed);
			return sw.Elapsed;
		}

		/// <summary>
		/// 测量有返回值方法的执行时间
		/// 自动记录性能数据到调试输出和可选的文件日志
		/// </summary>
		/// <typeparam name="T">返回值类型</typeparam>
		/// <param name="func">要执行的函数</param>
		/// <param name="callerMethod">方法名称（默认为调用者方法名）</param>
		/// <param name="filePath">调用者文件路径（默认为调用者文件路径）</param>
		/// <returns>元组：方法返回值和执行耗时</returns>
		public static (T result, TimeSpan elapsed) Time<T>(Func<T> func,[CallerMemberName] string callerMethod = "", [CallerFilePath] string filePath = "")
		{
			var sw = Stopwatch.StartNew();
			var result = func();
			sw.Stop();

			// 构建方法标识符，包含文件名和方法名
			string methodIdentifier = string.IsNullOrEmpty(filePath) 
				? callerMethod 
				: $"{Path.GetFileName(filePath)} | {callerMethod}";
			
			LogPerformance(methodIdentifier, sw.Elapsed);
			return (result, sw.Elapsed);
		}

		#endregion Public Methods

		#region Private Methods

		private static void FlushBuffer()
		{
			if(_writer == null)
			{
				try
				{
					// 延迟初始化写入器
					_writer = new StreamWriter(LogFilePath,append: true);
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
				while(_buffer.Count > 0)
				{
					_writer.WriteLine(_buffer.Dequeue());
				}
				_writer.Flush();
			} catch
			{
				// 写入失败时清理资源
				_writer.Dispose();
				_writer = null;
				_buffer.Clear();
			}
		}

		// 在 Profiler.cs 中
		private static void LogPerformance(string callerMethod,TimeSpan elapsed)
		{
			string message = $" {callerMethod} 执行耗时: {elapsed.TotalMilliseconds:F3} ms";

			// 始终输出到调试控制台
			Debug.WriteLine($"[性能监控]\t{message}");

			// 文件日志
			if(EnableFileLogging)
			{
				var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}]\t[性能监控]\t{callerMethod}\t执行耗时:{elapsed.TotalMilliseconds:F3} ms";

				lock(_lockObj)
				{
					_buffer.Enqueue(line);

					if(_buffer.Count>=BufferCapacity||_writer==null)
					{
						FlushBuffer();
					}
				}
			}
		}

		#endregion Private Methods
	}
}

namespace Project.Utilities.Extensions
{
	// 在Profiler类中添加扩展方法
	public static class ProfilerEx
	{
		#region Public Methods

		/// <summary>
		/// [扩展方法] 测量操作执行时间 - 使用流畅API风格
		/// </summary>
		/// <remarks>注意：此方法会使所有Action获得Time()方法</remarks>
		public static TimeSpan Time(
			this Action action,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string filePath = "")
		{
			return Profiler.Time(action,callerMethod,filePath);
		}

		/// <summary>
		/// [扩展方法] 测量操作执行时间 - 使用流畅API风格
		/// </summary>
		/// <remarks>注意：此方法会使所有Func获得Time()方法</remarks>
		public static (TResult Result, TimeSpan Elapsed) Time<TResult>(
			this Func<TResult> func,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string filePath = "")
		{
			return Profiler.Time(func,callerMethod,filePath);
		}

		// 可选：添加常用参数类型的重载
		public static TimeSpan Time<Targs>(
			this Action<Targs> action,
			Targs arg,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string filePath = "")
		{
			return Time(() => action(arg),callerMethod,filePath);
		}

		public static (TResult Result, TimeSpan Elapsed) Time<Targs, TResult>(
			this Func<Targs,TResult> func,
			Targs arg,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string filePath = "")
		{
			return Time(() => func(arg),callerMethod,filePath);
		}

		#endregion Public Methods
	}
}

/*
// =============================================
// 超级简单使用示例（放在代码文件末尾即可）
// =============================================

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

// 示例4：实际使用场景
// 在方法开始时测量性能
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

 */


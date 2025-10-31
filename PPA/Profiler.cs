using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;

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
		private static readonly Queue<string> _buffer = new Queue<string>(); // 日志缓冲区
		private static readonly object _lockObj = new object(); // 线程同步锁
		private static StreamWriter _writer; // 文件写入器

		#endregion Private Fields

		#region Public Methods

		/// <summary>
		/// 测量无返回值方法的执行时间
		/// 自动记录性能数据到调试输出和可选的文件日志
		/// </summary>
		/// <param name="action">要执行的操作</param>
		/// <param name="methodName">方法名称（默认为调用者方法名）</param>
		/// <returns>执行耗时</returns>
		public static TimeSpan Time(Action action,[CallerMemberName] string methodName = "")
		{
			var sw = Stopwatch.StartNew();
			try
			{
				action();
				return sw.Elapsed;
			} finally
			{
				sw.Stop();
				LogPerformance(methodName,sw.Elapsed);
			}
		}

		/// <summary>
		/// 测量带返回值方法的执行时间
		/// 自动记录性能数据到调试输出和可选的文件日志
		/// </summary>
		/// <typeparam name="T">返回值类型</typeparam>
		/// <param name="func">要执行的函数</param>
		/// <param name="methodName">方法名称（默认为调用者方法名）</param>
		/// <returns>包含结果和执行耗时的元组</returns>
		public static (T Result, TimeSpan Elapsed) Time<T>(Func<T> func,[CallerMemberName] string methodName = "")
		{
			var sw = Stopwatch.StartNew();
			try
			{
				var result = func();
				return (result, sw.Elapsed);
			} finally
			{
				sw.Stop();
				LogPerformance(methodName,sw.Elapsed);
			}
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

		private static void LogPerformance(string methodName,TimeSpan elapsed)
		{
			// 保持原有控制台日志功能
			Debug.WriteLine($"[性能监控]\t{methodName}\t执行耗时: {elapsed.TotalMilliseconds:F3} ms");

			// 文件日志使用缓冲写入
			if(EnableFileLogging)
			{
				var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}]\t[性能监控]\t{methodName}\t执行耗时:{elapsed.TotalMilliseconds:F3} ms";

				lock(_lockObj)
				{
					_buffer.Enqueue(line);

					// 缓冲区满或写入器未初始化时刷新
					if(_buffer.Count >= BufferCapacity || _writer == null)
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
			[CallerMemberName] string methodName = "",
			[CallerFilePath] string filePath = "")
		{
			return Profiler.Time(action,$"{Path.GetFileName(filePath)} | {methodName}");
		}

		/// <summary>
		/// [扩展方法] 测量操作执行时间 - 使用流畅API风格
		/// </summary>
		/// <remarks>注意：此方法会使所有Func获得Time()方法</remarks>
		public static (TResult Result, TimeSpan Elapsed) Time<TResult>(
			this Func<TResult> func,
			[CallerMemberName] string methodName = "",
			[CallerFilePath] string filePath = "")
		{
			return Profiler.Time(func,$"{Path.GetFileName(filePath)} | {methodName}");
		}

		// 可选：添加常用参数类型的重载
		public static TimeSpan Time<Targs>(
			this Action<Targs> action,
			Targs arg,
			[CallerMemberName] string methodName = "",
			[CallerFilePath] string filePath = "")
		{
			return Time(() => action(arg),filePath,methodName);
		}

		public static (TResult Result, TimeSpan Elapsed) Time<Targs, TResult>(
			this Func<Targs,TResult> func,
			Targs arg,
			[CallerMemberName] string methodName = "",
			[CallerFilePath] string filePath = "")
		{
			return Time(() => func(arg),filePath,methodName);
		}

		#endregion Public Methods
	}
}

/*
// 自动获取文件和方法名
ProfilerEx.Time(() => action(args));
elapsed = action.Time(callerFile,callerMethod);
elapsed = ProfilerEx.Time(action,callerFile,callerMethod);
// 手动指定位置（调试复杂调用链）
ProfilerEx.Time(() => action(args), "Service.cs | ProcessRequest");
elapsed = ProfilerEx.Time(action,$"{callerFile} | {callerMethod}");
(result, elapsed) = ProfilerEx.Time(func,$"{callerFile} | {callerMethod}");
// 推荐 - 使用编译器特性
ProfilerEx.Time(action);
// 自动扩展函数方法
(result, elapsed) = func.Time(callerFile,callerMethod);

实际使用建议
场景1：简单调用（推荐）
// 直接使用闭包 - 最简洁
Profiler.Time(() => repository.Save(user));
场景2：需要明确参数传递
// 显式传递参数 - 更明确
Profiler.Time((u) => repository.Save(u), user);
场景3：多个参数
// 多个参数使用闭包
Profiler.Time(() => service.Process(order, user));
// 或显式传递
Profiler.Time((o, u) => service.Process(o, u), order, user);
 */
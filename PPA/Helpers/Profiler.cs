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
        public static TimeSpan Time(Action action, [CallerMemberName] string callerMethod = "", [CallerFilePath] string filePath = "")
        {
            var sw = Stopwatch.StartNew();
            action();
            sw.Stop();

            // 直接调用LogMessage方法记录性能数据
            string message = $"执行耗时: {sw.Elapsed.TotalMilliseconds:F3} ms";
            LogMessage(message, "性能监控", callerMethod, filePath);
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
        public static (T result, TimeSpan elapsed) Time<T>(Func<T> func, [CallerMemberName] string callerMethod = "", [CallerFilePath] string filePath = "")
        {
            var sw = Stopwatch.StartNew();
            var result = func();
            sw.Stop();

            // 直接调用LogMessage方法记录性能数据
            string message = $"执行耗时: {sw.Elapsed.TotalMilliseconds:F3} ms";
            LogMessage(message, "性能监控", callerMethod, filePath);
            return (result, sw.Elapsed);
        }

        #endregion Public Methods

        #region Public Methods

        /// <summary>
        /// 记录自定义日志信息
        /// 开发状态下输出到Debug控制台，非开发状态下写入日志文件
        /// </summary>
        /// <param name="message">日志消息内容</param>
        /// <param name="logLevel">日志级别（如：INFO, WARN, ERROR等）</param>
        /// <param name="callerMethod">调用者方法名（自动获取）</param>
        /// <param name="filePath">调用者文件路径（自动获取）</param>
        public static void LogMessage(string message, string logLevel = "INFO", [CallerMemberName] string callerMethod = "", [CallerFilePath] string filePath = "")
        {
            // 构建日志标识符，包含文件名和方法名
            string methodIdentifier = string.IsNullOrEmpty(filePath)
                ? callerMethod
                : $"{Path.GetFileNameWithoutExtension(filePath)}.{callerMethod}";

#if DEBUG
            // 开发状态：输出到Debug控制台
            Debug.WriteLine($"[{logLevel}]\t{methodIdentifier}\t{message}");
#else
			// 非开发状态：写入日志文件
			var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}]\t[{logLevel}]\t{methodIdentifier}\t{message}";

			lock(_lockObj)
			{
				_buffer.Enqueue(line);

				if(_buffer.Count >= BufferCapacity || _writer == null)
				{
					FlushBuffer();
				}
			}
#endif
        }

        #endregion Public Methods

        #region Private Methods

        private static void FlushBuffer()
        {
            if (_writer == null)
            {
                try
                {
                    // 延迟初始化写入器
                    _writer = new StreamWriter(LogFilePath, append: true);
                }
                catch
                {
                    // 初始化失败时清空缓冲区
                    _buffer.Clear();
                    return;
                }
            }

            try
            {
                // 写入所有缓冲日志
                while (_buffer.Count > 0)
                {
                    _writer.WriteLine(_buffer.Dequeue());
                }
                _writer.Flush();
            }
            catch
            {
                // 写入失败时清理资源
                _writer.Dispose();
                _writer = null;
                _buffer.Clear();
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
            return Profiler.Time(action, callerMethod, filePath);
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
            return Profiler.Time(func, callerMethod, filePath);
        }

        // 可选：添加常用参数类型的重载
        public static TimeSpan Time<Targs>(
            this Action<Targs> action,
            Targs arg,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string filePath = "")
        {
            return Time(() => action(arg), callerMethod, filePath);
        }

        public static (TResult Result, TimeSpan Elapsed) Time<Targs, TResult>(
            this Func<Targs, TResult> func,
            Targs arg,
            [CallerMemberName] string callerMethod = "",
            [CallerFilePath] string filePath = "")
        {
            return Time(() => func(arg), callerMethod, filePath);
        }

        /// <summary>
        /// [扩展方法] 记录自定义日志信息
        /// 开发状态下输出到Debug控制台，非开发状态下写入日志文件
        /// </summary>
        /// <param name="obj">任意对象（仅用于扩展方法语法）</param>
        /// <param name="message">日志消息内容</param>
        /// <param name="logLevel">日志级别（如：INFO, WARN, ERROR等）</param>
        /// <param name="callerMethod">调用者方法名（自动获取）</param>
        /// <param name="filePath">调用者文件路径（自动获取）</param>
        public static void Log(this object obj, string message, string logLevel = "INFO", [CallerMemberName] string callerMethod = "", [CallerFilePath] string filePath = "")
        {
            Profiler.LogMessage(message, logLevel, callerMethod, filePath);
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


// 直接调用静态方法
Profiler.LogMessage("这是一条信息日志");
Profiler.LogMessage("这是一条错误日志", "ERROR");

// 使用扩展方法
var myObject = new SomeClass();
myObject.Log("这是通过扩展方法记录的日志");
myObject.Log("这是一条警告日志", "WARN");

 */


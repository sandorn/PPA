using System;
using System.Runtime.CompilerServices;

namespace PPA.Core.Abstraction.Infrastructure
{
	/// <summary>
	/// 统一的日志接口 提供标准化的日志记录功能，方便替换底层实现
	/// </summary>
	/// <remarks>
	/// 此接口定义了统一的日志记录方法，支持以下日志级别：
	/// <list type="bullet">
	/// <item>
	/// <description> <see cref="LogInformation(string, string, string)" /> - 信息日志 </description>
	/// </item>
	/// <item>
	/// <description> <see cref="LogWarning(string, string, string)" /> - 警告日志 </description>
	/// </item>
	/// <item>
	/// <description> <see cref="LogError(string, Exception, string, string)" /> - 错误日志 </description>
	/// </item>
	/// <item>
	/// <description> <see cref="LogDebug(string, string, string)" /> - 调试日志 </description>
	/// </item>
	/// </list>
	/// 所有方法都支持自动获取调用位置信息（通过 <c> [CallerMemberName] </c> 和 <c> [CallerFilePath] </c> 特性），
	/// 无需手动传递调用方法名和文件路径。 默认实现为 <see cref="PPA.Core.Logging.ProfilerLoggerAdapter" />，基于现有的
	/// <see cref="PPA.Core.Profiler" /> 系统。
	/// </remarks>
	public interface ILogger
	{
		/// <summary>
		/// 记录指定级别的日志
		/// </summary>
		/// <param name="level"> 日志级别 </param>
		/// <param name="message"> 日志消息 </param>
		/// <param name="exception"> 可选的异常对象 </param>
		/// <param name="callerMethod"> 调用方法名（自动获取，无需手动传递） </param>
		/// <param name="callerFile"> 调用文件路径（自动获取，无需手动传递） </param>
		void Log(LogLevel level,string message,Exception exception = null,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "");

		/// <summary>
		/// 记录信息级别的日志
		/// </summary>
		/// <param name="message"> 日志消息 </param>
		/// <param name="callerMethod"> 调用方法名（自动获取，无需手动传递） </param>
		/// <param name="callerFile"> 调用文件路径（自动获取，无需手动传递） </param>
		void LogInformation(string message,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "");

		/// <summary>
		/// 记录警告级别的日志
		/// </summary>
		/// <param name="message"> 日志消息 </param>
		/// <param name="callerMethod"> 调用方法名（自动获取，无需手动传递） </param>
		/// <param name="callerFile"> 调用文件路径（自动获取，无需手动传递） </param>
		void LogWarning(string message,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "");

		/// <summary>
		/// 记录调试级别的日志
		/// </summary>
		/// <param name="message"> 日志消息 </param>
		/// <param name="callerMethod"> 调用方法名（自动获取，无需手动传递） </param>
		/// <param name="callerFile"> 调用文件路径（自动获取，无需手动传递） </param>
		void LogDebug(string message,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "");

		/// <summary>
		/// 记录错误级别的日志
		/// </summary>
		/// <param name="message"> 日志消息 </param>
		/// <param name="exception"> 可选的异常对象 </param>
		/// <param name="callerMethod"> 调用方法名（自动获取，无需手动传递） </param>
		/// <param name="callerFile"> 调用文件路径（自动获取，无需手动传递） </param>
		void LogError(string message,Exception exception = null,
			[CallerMemberName] string callerMethod = "",
			[CallerFilePath] string callerFile = "");
	}
}

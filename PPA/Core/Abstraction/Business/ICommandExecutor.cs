using System;
using System.Collections.Generic;
using PPA.Core.Abstraction.Presentation;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// Office 原生命令执行器接口
	/// 用于执行 PowerPoint 内置菜单命令和功能区命令
	/// </summary>
	public interface ICommandExecutor
	{
		/// <summary>
		/// 通过 MSO 命令名称执行命令（推荐方式）
		/// </summary>
		/// <param name="msoCommandName">MSO 命令名称，例如 "Paste", "Copy", "Bold"</param>
		/// <returns>是否执行成功</returns>
		bool ExecuteMso(string msoCommandName);

		/// <summary>
		/// 执行命令并返回详细结果
		/// </summary>
		/// <remarks>
		/// 使用原生 COM 对象的 FindControl 方法查找命令。如果 FindControl 失败，则返回失败结果。
		/// 传统方式，保留，未使用。
		/// </remarks>
		/// <param name="commandId">命令 ID</param>
		/// <returns>命令执行结果详情</returns>
		CommandExecutionResult ExecuteCommandById(int commandId);

		/// <summary>
		/// 通过菜单路径执行命令（例如 "File|Save As"）
		/// </summary>
		/// <param name="menuPath">菜单路径，使用 "|" 分隔层级</param>
		/// <returns>是否执行成功</returns>
		bool ExecuteMenuPath(string menuPath);

	}

	/// <summary>
	/// 命令执行结果详情
	/// </summary>
	public class CommandExecutionResult
	{
		/// <summary>
		/// 命令 ID
		/// </summary>
		public int CommandId { get; set; }

		/// <summary>
		/// 是否执行成功
		/// </summary>
		public bool Success { get; set; }

		/// <summary>
		/// 是否找到控件
		/// </summary>
		public bool ControlFound { get; set; }

		/// <summary>
		/// 控件标题
		/// </summary>
		public string ControlCaption { get; set; }

		/// <summary>
		/// 控件类型
		/// </summary>
		public string ControlType { get; set; }

		/// <summary>
		/// 控件是否可用
		/// </summary>
		public bool IsEnabled { get; set; }

		/// <summary>
		/// 执行时间
		/// </summary>
		public DateTime? ExecutionTime { get; set; }

		/// <summary>
		/// 错误消息
		/// </summary>
		public string ErrorMessage { get; set; }

		/// <summary>
		/// 异常信息
		/// </summary>
		public Exception Exception { get; set; }
	}
}


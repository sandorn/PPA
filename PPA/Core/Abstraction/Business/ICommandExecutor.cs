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
		/// 通过命令 ID 执行命令
		/// </summary>
		/// <param name="commandId">命令 ID</param>
		/// <returns>是否执行成功</returns>
		bool ExecuteCommandById(int commandId);

		/// <summary>
		/// 通过菜单路径执行命令（例如 "File|Save As"）
		/// </summary>
		/// <param name="menuPath">菜单路径，使用 "|" 分隔层级</param>
		/// <returns>是否执行成功</returns>
		bool ExecuteMenuPath(string menuPath);

	}
}


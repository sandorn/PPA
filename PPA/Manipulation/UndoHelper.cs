using PPA.Core;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Logging;
using PPA.Utilities;
using System;
using System.Threading.Tasks;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Manipulation
{
	/// <summary>
	/// 撤销操作助手类 提供统一的撤销/重做管理，支持描述性撤销名称和撤销组
	/// </summary>
	/// <remarks>
	/// 此类封装了 PowerPoint 的撤销/重做功能，提供以下功能：
	/// <list type="bullet">
	/// <item>
	/// <description> 开始新的撤销单元 - <see cref="BeginUndoEntry(NETOP.Application, string)" /> </description>
	/// </item>
	/// <item>
	/// <description> 在撤销组中执行操作 - <see cref="ExecuteInUndoGroup(NETOP.Application, string, Action)" /> </description>
	/// </item>
	/// <item>
	/// <description> 预定义的撤销操作名称 - <see cref="UndoNames" /> </description>
	/// </item>
	/// </list>
	/// 注意：PowerPoint API 不支持设置撤销操作的显示名称，此处的 undoName 参数仅用于日志记录。
	/// </remarks>
	public static class UndoHelper
	{
		private static readonly ILogger _logger = LoggerProvider.GetLogger();
		private static ILogger Logger => _logger??LoggerProvider.GetLogger();

		#region Public Methods

		/// <summary>
		/// 开始一个新的撤销单元（同步版本）
		/// </summary>
		/// <param name="netApp"> PowerPoint 应用程序实例，如果为 null 则不执行任何操作 </param>
		/// <param name="undoName"> 撤销操作的名称（仅用于日志记录，PowerPoint API 不支持设置撤销名称） </param>
		/// <remarks>
		/// 此方法会调用 PowerPoint 的 <c> StartNewUndoEntry() </c> 方法，将后续操作合并为一个撤销单元。 如果无法获取有效的
		/// Application 对象，会记录警告日志但不抛出异常，避免影响主流程。
		/// </remarks>
		public static void BeginUndoEntry(NETOP.Application netApp,string undoName = null)
		{
			if(netApp==null) return;

			var safeApp = ApplicationHelper.EnsureValidNetApplication(netApp);
			if(safeApp==null)
			{
				_logger.LogWarning("无法获取有效的 Application，跳过撤销单元");
				return;
			}

			try
			{
				safeApp.StartNewUndoEntry();

				// 记录撤销操作（用于日志追踪）
				if(!string.IsNullOrEmpty(undoName))
				{
					_logger.LogInformation($"开始撤销单元: {undoName}");
				}
			} catch(Exception ex)
			{
				_logger.LogWarning($"创建撤销单元失败: {ex.Message}");
				// 不抛出异常，避免影响主流程
			}
		}

		/// <summary>
		/// 开始一个新的撤销单元（异步版本）
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例，如果为 null 则不执行任何操作 </param>
		/// <param name="undoName"> 撤销操作的名称（仅用于日志记录，PowerPoint API 不支持设置撤销名称） </param>
		/// <returns> 表示异步操作的 Task </returns>
		/// <remarks>
		/// 此方法会在 UI 线程上执行 <see cref="BeginUndoEntry(NETOP.Application, string)" /> 方法。 必须在 UI 线程上调用，否则可能抛出异常。
		/// </remarks>
		public static async Task BeginUndoEntryAsync(NETOP.Application app,string undoName = null)
		{
			if(app==null) return;

			await AsyncOperationHelper.RunOnUIThread(() =>
			{
				BeginUndoEntry(app,undoName);
			});
		}

		/// <summary>
		/// 在撤销组中执行操作（将多个操作合并为一个撤销单元）
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例，如果为 null 则不执行任何操作 </param>
		/// <param name="undoName"> 撤销组的名称，用于日志记录 </param>
		/// <param name="action"> 要执行的操作，如果为 null 则不执行任何操作 </param>
		/// <remarks>
		/// 此方法会先调用 <see cref="BeginUndoEntry(NETOP.Application, string)" /> 开始新的撤销单元， 然后执行
		/// action，最后将 action 中的所有操作合并为一个撤销单元。 如果 action 中抛出异常，异常会被重新抛出，让上层处理。
		/// </remarks>
		/// <example>
		/// <code>
		///UndoHelper.ExecuteInUndoGroup(app, "批量美化", () =&gt;
		///{
		///TableFormatHelper.FormatTables(table1);
		///TableFormatHelper.FormatTables(table2);
		///TableFormatHelper.FormatTables(table3);
		///});
		/// </code>
		/// </example>
		public static void ExecuteInUndoGroup(NETOP.Application app,string undoName,Action action)
		{
			if(app==null||action==null) return;

			BeginUndoEntry(app,undoName);
			try
			{
				action();
			} catch
			{
				throw; // 重新抛出异常，让上层处理
			}
		}

		/// <summary>
		/// 在撤销组中执行异步操作（将多个操作合并为一个撤销单元）
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例，如果为 null 则不执行任何操作 </param>
		/// <param name="undoName"> 撤销组的名称，用于日志记录 </param>
		/// <param name="asyncAction"> 要执行的异步操作，如果为 null 则不执行任何操作 </param>
		/// <returns> 表示异步操作的 Task </returns>
		/// <remarks>
		/// 此方法会先调用 <see cref="BeginUndoEntryAsync(NETOP.Application, string)" /> 开始新的撤销单元， 然后执行
		/// asyncAction，最后将 asyncAction 中的所有操作合并为一个撤销单元。 如果 asyncAction 中抛出异常，异常会被重新抛出，让上层处理。
		/// </remarks>
		public static async Task ExecuteInUndoGroupAsync(NETOP.Application app,string undoName,Func<Task> asyncAction)
		{
			if(app==null||asyncAction==null) return;

			await BeginUndoEntryAsync(app,undoName);
			try
			{
				await asyncAction();
			} catch
			{
				throw; // 重新抛出异常，让上层处理
			}
		}

		#endregion Public Methods

		#region Predefined Undo Names

		/// <summary>
		/// 预定义的撤销操作名称（使用本地化字符串）
		/// </summary>
		/// <remarks> 此类提供常用的撤销操作名称，使用 <see cref="ResourceManager" /> 进行本地化。 这些名称用于日志记录，帮助追踪用户操作。 </remarks>
		public static class UndoNames
		{
			/// <summary>
			/// 美化表格操作的撤销名称
			/// </summary>
			public static string FormatTables => ResourceManager.GetString("Undo_FormatTables","美化表格");

			/// <summary>
			/// 美化文本操作的撤销名称
			/// </summary>
			public static string FormatText => ResourceManager.GetString("Undo_FormatText","美化文本");

			/// <summary>
			/// 美化图表操作的撤销名称
			/// </summary>
			public static string FormatCharts => ResourceManager.GetString("Undo_FormatCharts","美化图表");

			/// <summary>
			/// 对齐形状操作的撤销名称
			/// </summary>
			public static string AlignShapes => ResourceManager.GetString("Undo_AlignShapes","对齐形状");

			/// <summary>
			/// 创建外框操作的撤销名称
			/// </summary>
			public static string CreateBoundingBox => ResourceManager.GetString("Undo_CreateBoundingBox","创建外框");

			/// <summary>
			/// 隐藏对象操作的撤销名称
			/// </summary>
			public static string HideShapes => ResourceManager.GetString("Undo_HideShapes","隐藏对象");

			/// <summary>
			/// 显示对象操作的撤销名称
			/// </summary>
			public static string ShowShapes => ResourceManager.GetString("Undo_ShowShapes","显示对象");
		}

		#endregion Predefined Undo Names
	}
}

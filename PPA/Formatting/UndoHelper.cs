using PPA.Core;
using PPA.Utilities;
using System;
using System.Threading.Tasks;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 撤销操作助手类 提供统一的撤销/重做管理，支持描述性撤销名称和撤销组
	/// </summary>
	public static class UndoHelper
	{
		#region Public Methods

		/// <summary>
		/// 开始一个新的撤销单元（同步版本）
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		/// <param name="undoName"> 撤销操作的名称（仅用于日志记录，PowerPoint API 不支持设置撤销名称） </param>
		public static void BeginUndoEntry(NETOP.Application app,string undoName = null)
		{
			if(app==null) return;

			try
			{
				// 优先使用原生 PowerPoint Application 对象（性能更好）
				var nativeApp = Globals.ThisAddIn?.NativeApp;

				if(nativeApp!=null)
				{
					nativeApp.StartNewUndoEntry();
				} else
				{
					// 回退到 NetOffice 对象
					app.StartNewUndoEntry();
				}

				// 记录撤销操作（用于日志追踪）
				if(!string.IsNullOrEmpty(undoName))
				{
					Profiler.LogMessage($"开始撤销单元: {undoName}","INFO");
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"创建撤销单元失败: {ex.Message}","WARN");
				// 不抛出异常，避免影响主流程
			}
		}

		/// <summary>
		/// 开始一个新的撤销单元（异步版本） 必须在 UI 线程上调用
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		/// <param name="undoName"> 撤销操作的名称（仅用于日志记录，PowerPoint API 不支持设置撤销名称） </param>
		/// <returns> 表示异步操作的 Task </returns>
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
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		/// <param name="undoName"> 撤销组的名称 </param>
		/// <param name="action"> 要执行的操作 </param>
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
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		/// <param name="undoName"> 撤销组的名称 </param>
		/// <param name="asyncAction"> 要执行的异步操作 </param>
		/// <returns> 表示异步操作的 Task </returns>
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
		public static class UndoNames
		{
			public static string FormatTables => ResourceManager.GetString("Undo_FormatTables","美化表格");
			public static string FormatText => ResourceManager.GetString("Undo_FormatText","美化文本");
			public static string FormatCharts => ResourceManager.GetString("Undo_FormatCharts","美化图表");
			public static string AlignShapes => ResourceManager.GetString("Undo_AlignShapes","对齐形状");
			public static string CreateBoundingBox => ResourceManager.GetString("Undo_CreateBoundingBox","创建外框");
			public static string HideShapes => ResourceManager.GetString("Undo_HideShapes","隐藏对象");
			public static string ShowShapes => ResourceManager.GetString("Undo_ShowShapes","显示对象");
		}

		#endregion Predefined Undo Names
	}
}

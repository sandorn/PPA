using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Shape;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 表格批量操作辅助类
	/// </summary>
	public static class TableBatchHelper
	{
		/// <summary>
		/// 同步美化表格（原始版本，保留用于向后兼容）
		/// </summary>
		/// <remarks>
		/// <para>
		/// 注意：此方法为同步版本，在执行耗时操作时会阻塞 PowerPoint UI 线程。 建议使用异步版本 <see cref="Bt501_ClickAsync" /> 以获得更好的用户体验。
		/// </para>
		/// <para>
		/// 保留此方法的原因：
		/// 1. 向后兼容：确保现有代码调用不受影响
		/// 2. 简单场景：对于少量表格，同步执行可能更简单
		/// 3. 调试方便：同步代码更容易调试和追踪
		/// </para>
		/// </remarks>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		public static void Bt501_Click(NETOP.Application app)
		{
			UndoHelper.BeginUndoEntry(app,UndoHelper.UndoNames.FormatTables);
			var slide = ShapeUtils.TryGetCurrentSlide(app);

			ExHandler.Run(() =>
			{
				var sel = ShapeUtils.ValidateSelection(app);
				int tableCount = 0;

				if(sel!=null)
				{
					// 有选中对象的情况，美化选中对象中的表格
					if(sel is NETOP.Shape shape)
					{
						if(shape.HasTable==MsoTriState.msoTrue)
						{
							TableFormatHelper.FormatTables(shape.Table);
							tableCount++;
						}
					} else if(sel is NETOP.ShapeRange shapes)
					{
						foreach(NETOP.Shape s in shapes)
						{
							if(s.HasTable==MsoTriState.msoTrue)
							{
								TableFormatHelper.FormatTables(s.Table);
								tableCount++;
							}
						}
					}

					if(tableCount>0)
						Toast.Show(ResourceManager.GetString("Toast_FormatTables_Success","成功美化 {0} 个表格",tableCount),Toast.ToastType.Success);
					else
						Toast.Show(ResourceManager.GetString("Toast_FormatTables_NoSelection","选中的对象中没有表格"),Toast.ToastType.Info);
				} else
				{
					// 未选中对象的情况，美化当前幻灯片所有表格
					if(slide!=null)
					{
						foreach(NETOP.Shape shape in slide.Shapes)
						{
							if(shape.HasTable==MsoTriState.msoTrue)
							{
								TableFormatHelper.FormatTables(shape.Table);
								tableCount++;
							}
						}

						if(tableCount>0)
							Toast.Show(ResourceManager.GetString("Toast_FormatTables_Success","成功美化 {0} 个表格",tableCount),Toast.ToastType.Success);
						else
							Toast.Show(ResourceManager.GetString("Toast_FormatTables_NoTables","当前幻灯片上没有表格"),Toast.ToastType.Info);
					}
				}
			},enableTiming: true);
		}

		/// <summary>
		/// 异步美化表格（支持进度报告和取消）
		/// </summary>
		/// <remarks>
		/// <para>
		/// 这是 <see cref="Bt501_Click" /> 的异步版本，提供以下改进：
		/// 1. 非阻塞执行：不会冻结 PowerPoint UI 线程
		/// 2. 进度反馈：通过 <paramref name="progress" /> 参数报告美化进度
		/// 3. 取消支持：通过 <paramref name="cancellationToken" /> 支持取消操作
		/// 4. 更好的用户体验：执行耗时操作时用户可以继续使用 PowerPoint
		/// </para>
		/// <para>
		/// 使用场景：
		/// - 美化大量表格时（推荐使用此异步版本）
		/// - 需要进度反馈时
		/// - 需要支持取消操作时
		/// </para>
		/// <para> 注意：所有 Office COM 对象操作都在 UI 线程执行，确保线程安全。 </para>
		/// </remarks>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		/// <param name="progress"> 进度报告对象，用于报告美化进度（可选） </param>
		/// <param name="cancellationToken"> 取消令牌，用于取消正在进行的操作（可选） </param>
		/// <returns> 表示异步操作的 Task </returns>
		public static async Task Bt501_ClickAsync(
			NETOP.Application app,
			IProgress<AsyncProgress> progress = null,
			CancellationToken cancellationToken = default)
		{
			try
			{
				// 必须在 UI 线程执行撤销操作
				await UndoHelper.BeginUndoEntryAsync(app,UndoHelper.UndoNames.FormatTables);

				var slide = await AsyncOperationHelper.RunOnUIThread(() =>
				{
					return ShapeUtils.TryGetCurrentSlide(app);
				});

				if(slide==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				// 在 UI 线程收集表格信息
				var tables = await AsyncOperationHelper.RunOnUIThread(() =>
				{
					var sel = ShapeUtils.ValidateSelection(app);
					var result = new List<(NETOP.Shape shape, NETOP.Table table)>();

					if (sel != null)
					{
                        // 有选中对象的情况，收集选中对象中的表格
                        if (sel is NETOP.Shape shape && shape.HasTable == MsoTriState.msoTrue)
						{
							result.Add((shape, shape.Table));
							progress?.Report(new AsyncProgress(10, ResourceManager.GetString("Progress_TableFound", "发现表格"), 1, 1));
						}
						else if (sel is NETOP.ShapeRange shapes)
						{
							int count = 0;
							foreach (NETOP.Shape s in shapes)
							{
								if (s.HasTable == MsoTriState.msoTrue)
								{
									result.Add((s, s.Table));
									count++;
									progress?.Report(new AsyncProgress(
										10,
										ResourceManager.GetString("Progress_TableFound_Count", "发现表格 {0}", count),
										count,
										shapes.Count));
								}
							}
						}
					}
					else
					{
                        // 未选中对象的情况，收集当前幻灯片所有表格
                        foreach (NETOP.Shape shape in slide.Shapes)
						{
							if (shape.HasTable == MsoTriState.msoTrue)
							{
								result.Add((shape, shape.Table));
							}
						}
					}

					return result;
				});

				cancellationToken.ThrowIfCancellationRequested();

				int total = tables.Count;

				// 如果没有找到表格
				if(total==0)
				{
					var hasSelection = await AsyncOperationHelper.RunOnUIThread(() =>
					{
						return ShapeUtils.ValidateSelection(app) != null;
					});
					Toast.Show(hasSelection ? ResourceManager.GetString("Toast_FormatTables_NoSelection","选中的对象中没有表格") : ResourceManager.GetString("Toast_FormatTables_NoTables","当前幻灯片上没有表格"),Toast.ToastType.Info);
					return;
				}

				// 报告找到的表格数量（简化日志，不逐条扫描）
				progress?.Report(new AsyncProgress(10,ResourceManager.GetString("Progress_TablesFound","发现 {0} 个表格",total),total,total));

				progress?.Report(new AsyncProgress(20,ResourceManager.GetString("Progress_FormatTables_Start","开始美化 {0} 个表格",total),0,total));

				// 在 UI 线程逐个美化表格（同步执行，但允许 UI 更新）
				for(int i = 0;i<total;i++)
				{
					cancellationToken.ThrowIfCancellationRequested();

					NETOP.Shape shape = tables[i].shape;
					NETOP.Table table = tables[i].table;

					// 计算当前表格的进度（20% 开始，80% 结束）
					int currentProgress = 20 + (int)((i * 60.0) / total);
					progress?.Report(new AsyncProgress(
						currentProgress,
						ResourceManager.GetString("Progress_FormatTable_Progress","美化表格 {0}/{1}",i+1,total),
						i+1,
						total));

					// 在 UI 线程执行美化
					await AsyncOperationHelper.RunOnUIThread(() =>
					{
						TableFormatHelper.FormatTables(table);
					});

					// 允许 UI 更新（每处理一个表格后）
					await Task.Delay(10,cancellationToken);
				}

				progress?.Report(new AsyncProgress(100,ResourceManager.GetString("Progress_FormatTables_Complete","成功美化 {0} 个表格",total),total,total));
			} catch
			{
				throw; // 重新抛出异常，让上层处理
			}
		}
	}
}

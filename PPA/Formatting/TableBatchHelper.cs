using Microsoft.Extensions.DependencyInjection;
using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Shape;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 表格批量操作辅助类
	/// </summary>
	internal class TableBatchHelper(ITableFormatHelper tableFormatHelper,IShapeHelper shapeHelper): ITableBatchHelper
	{
		private readonly ITableFormatHelper _tableFormatHelper = tableFormatHelper ?? throw new ArgumentNullException(nameof(tableFormatHelper));
		private readonly IShapeHelper _shapeHelper = shapeHelper ?? throw new ArgumentNullException(nameof(shapeHelper));

		#region ITableBatchHelper 实现

		public void FormatTables(NETOP.Application app)
		{
			if(app==null) throw new ArgumentNullException(nameof(app));
			FormatTablesInternal(app,_tableFormatHelper);
		}

		public Task FormatTablesAsync(
			NETOP.Application app,
			IProgress<AsyncProgress> progress = null,
			CancellationToken cancellationToken = default)
		{
			if(app==null) throw new ArgumentNullException(nameof(app));
			return FormatTablesInternalAsync(app,progress,cancellationToken,_tableFormatHelper);
		}

		#endregion

		#region 向后兼容的静态方法

		/// <summary>
		/// 同步美化表格（向后兼容入口）
		/// </summary>
		public static void Bt501_Click(NETOP.Application app, ITableFormatHelper tableFormatHelper = null)
		{
			var batchHelper = ResolveInstance(tableFormatHelper);
			batchHelper.FormatTables(app);
		}

		/// <summary>
		/// 异步美化表格（向后兼容入口）
		/// </summary>
		public static Task Bt501_ClickAsync(
			NETOP.Application app,
			IProgress<AsyncProgress> progress = null,
			CancellationToken cancellationToken = default,
			ITableFormatHelper tableFormatHelper = null)
		{
			var batchHelper = ResolveInstance(tableFormatHelper);
			return batchHelper.FormatTablesAsync(app,progress,cancellationToken);
		}

		private static ITableBatchHelper ResolveInstance(ITableFormatHelper overrideHelper)
		{
			if(overrideHelper!=null)
				return new TableBatchHelper(overrideHelper,ResolveShapeHelper());

			var serviceProvider = Globals.ThisAddIn?.ServiceProvider;
			if(serviceProvider!=null)
			{
				var resolved = serviceProvider.GetService(typeof(ITableBatchHelper)) as ITableBatchHelper;
				if(resolved!=null) return resolved;

				var tableHelper = serviceProvider.GetService(typeof(ITableFormatHelper)) as ITableFormatHelper;
				var shapeHelper = serviceProvider.GetService(typeof(IShapeHelper)) as IShapeHelper;
				if(tableHelper!=null&&shapeHelper!=null)
					return new TableBatchHelper(tableHelper,shapeHelper);
			}

			return new TableBatchHelper(
				new TableFormatHelper(FormattingConfig.Instance),
				ResolveShapeHelper());
		}

		private static IShapeHelper ResolveShapeHelper()
		{
			var serviceProvider = Globals.ThisAddIn?.ServiceProvider;
			var helper = serviceProvider?.GetService(typeof(IShapeHelper)) as IShapeHelper;
			return helper ?? ShapeUtils.Default;
		}

		#endregion

		#region 内部实现

		private void FormatTablesInternal(NETOP.Application app,ITableFormatHelper tableFormatHelper)
		{
			PPA.Core.Profiler.LogMessage($"FormatTablesInternal 开始，app类型={app?.GetType().Name ?? "null"}", "INFO");
			if(tableFormatHelper==null)
				throw new InvalidOperationException("无法获取 ITableFormatHelper 服务");

			UndoHelper.BeginUndoEntry(app,UndoHelper.UndoNames.FormatTables);

			var slide = _shapeHelper.TryGetCurrentSlide(app);
			PPA.Core.Profiler.LogMessage($"TryGetCurrentSlide 返回: {slide?.GetType().Name ?? "null"}", "INFO");

			ExHandler.Run(() =>
			{
				var sel = _shapeHelper.ValidateSelection(app);
				PPA.Core.Profiler.LogMessage($"ValidateSelection 返回: {sel?.GetType().Name ?? "null"}", "INFO");
				int tableCount = 0;
				var processedKeys = new List<object>();

				bool AlreadyProcessed(object key)
				{
					if(key == null) return false;
					foreach(var existing in processedKeys)
					{
						if(ReferenceEquals(existing, key))
							return true;
					}
					processedKeys.Add(key);
					return false;
				}

				bool ProcessAbstractShape(IShape abstractShape)
				{
					if(abstractShape == null) return false;

					string shapeName = null;
					try { shapeName = abstractShape.Name; } catch { }
					PPA.Core.Profiler.LogMessage($"处理抽象形状: {shapeName ?? "未知"}, HasTable={abstractShape.HasTable}", "INFO");

					if(!abstractShape.HasTable)
					{
						PPA.Core.Profiler.LogMessage($"抽象形状 {shapeName ?? "未知"} 不包含表格，跳过", "INFO");
						return false;
					}

					var key = (abstractShape as IComWrapper)?.NativeObject ?? abstractShape;
					if(AlreadyProcessed(key))
					{
						PPA.Core.Profiler.LogMessage($"抽象形状 {shapeName ?? "未知"} 已处理，跳过", "INFO");
						return false;
					}

					var table = abstractShape.GetTable();
					if(table != null)
					{
						PPA.Core.Profiler.LogMessage($"抽象形状 {shapeName ?? "未知"} 返回表格实例: {table.GetType().Name}", "INFO");
						tableFormatHelper.FormatTables(table);
						tableCount++;
						return true;
					}

					PPA.Core.Profiler.LogMessage($"抽象形状 {shapeName ?? "未知"} GetTable 返回 null", "WARN");
					return false;
				}


				if(sel!=null)
				{
					// 有选中对象的情况，美化选中对象中的表格
					if(sel is NETOP.Shape shape)
					{
						if(AlreadyProcessed(shape))
						{
							PPA.Core.Profiler.LogMessage($"形状 {shape.Name} 已经处理，跳过", "INFO");
						}
						else
						{
						// 检查是否是表格：先尝试 HasTable，如果失败则直接检查 Table 属性
						bool isTable = false;
						dynamic dynTable = null;
						try
						{
							PPA.Core.Profiler.LogMessage($"选中单个形状，HasTable={shape.HasTable}", "INFO");
							if(shape.HasTable == MsoTriState.msoTrue)
							{
								isTable = true;
								dynTable = shape.Table;
							}
						}
						catch
						{
							// HasTable 不可用，尝试直接检查 Table 属性
							PPA.Core.Profiler.LogMessage("HasTable 不可用，尝试直接检查 Table 属性", "INFO");
							try
							{
								dynamic dynShape = shape;
								dynTable = SafeGet(() => dynShape.Table, null);
								isTable = (dynTable != null);
								PPA.Core.Profiler.LogMessage($"直接检查 Table 属性: isTable={isTable}", "INFO");
							}
							catch(System.Exception ex)
							{
								PPA.Core.Profiler.LogMessage($"检查 Table 属性失败: {ex.Message}", "WARN");
							}
						}
						
						if(isTable && dynTable != null)
						{
							PPA.Core.Profiler.LogMessage("开始包装表格", "INFO");
							var iTable = AdapterUtils.WrapTable(app, shape, dynTable);
							PPA.Core.Profiler.LogMessage($"WrapTable 返回: {iTable?.GetType().Name ?? "null"}", "INFO");
							if(iTable != null)
							{
								tableFormatHelper.FormatTables(iTable);
								tableCount++;
							}
							else
							{
								PPA.Core.Profiler.LogMessage("WrapTable 返回 null，无法格式化", "ERROR");
							}
						}
						else
						{
							PPA.Core.Profiler.LogMessage("形状不是表格或 Table 属性为 null", "INFO");
						}
						}
					} else if(sel is NETOP.ShapeRange shapes)
					{
						PPA.Core.Profiler.LogMessage($"选中多个形状，Count={shapes.Count}", "INFO");
						try
						{
							foreach(NETOP.Shape s in shapes)
							{
								// 检查是否是表格：先尝试 HasTable，如果失败则直接检查 Table 属性
								bool isTable = false;
								dynamic dynTable = null;
								try
								{
									if(s.HasTable == MsoTriState.msoTrue)
									{
										isTable = true;
										dynTable = s.Table;
									}
								}
								catch
								{
									// HasTable 不可用，尝试直接检查 Table 属性
									try
									{
										dynamic dynShape = s;
										dynTable = SafeGet(() => dynShape.Table, null);
										isTable = (dynTable != null);
									}
									catch { }
								}
								
								if(isTable && dynTable != null)
								{
									if(AlreadyProcessed(s))
									{
										PPA.Core.Profiler.LogMessage($"形状 {s.Name} 已处理，跳过", "INFO");
										continue;
									}
									PPA.Core.Profiler.LogMessage($"处理形状 {s.Name}，检测到表格", "INFO");
									var iTable = AdapterUtils.WrapTable(app, s, dynTable);
									PPA.Core.Profiler.LogMessage($"WrapTable 返回: {iTable?.GetType().Name ?? "null"}", "INFO");
									if(iTable != null)
									{
										tableFormatHelper.FormatTables(iTable);
										tableCount++;
									}
									else
									{
										PPA.Core.Profiler.LogMessage("WrapTable 返回 null，无法格式化", "ERROR");
									}
								}
							}
						}
						catch(System.Exception ex)
						{
							// NetOffice 无法枚举 WPS ShapeRange，使用 dynamic 访问
							PPA.Core.Profiler.LogMessage($"NetOffice 枚举 ShapeRange 失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
							try
							{
								dynamic dynShapeRange = shapes;
								int rangeCount = SafeGet(() => (int)dynShapeRange.Count, 0);
								PPA.Core.Profiler.LogMessage($"使用 dynamic 访问 ShapeRange，Count={rangeCount}", "INFO");
								for(int i = 1; i <= rangeCount; i++)
								{
										dynamic dynShape = SafeGet(() => dynShapeRange[i], null);
										if(dynShape != null)
										{
											// WPS 中 HasTable 可能不可用，直接检查 Table 属性
											dynamic dynTable = null;
											bool hasTable = false;
											try
											{
												// 方式1：尝试通过 HasTable 属性
												hasTable = SafeGet(() => (bool)(dynShape.HasTable ?? false), false);
												if(hasTable)
												{
													dynTable = SafeGet(() => dynShape.Table, null);
												}
											}
											catch { }
											
											// 方式2：如果方式1失败，直接检查 Table 属性
											if(!hasTable || dynTable == null)
											{
												try
												{
													dynTable = SafeGet(() => dynShape.Table, null);
													hasTable = (dynTable != null);
												}
												catch { }
											}
											
											if(hasTable && dynTable != null)
											{
											if(AlreadyProcessed(dynShape))
											{
												PPA.Core.Profiler.LogMessage($"dynamic 形状 {i} 已处理，跳过", "INFO");
												continue;
											}
												PPA.Core.Profiler.LogMessage($"发现表格形状 {i}", "INFO");
												var iTable = AdapterUtils.WrapTable(app, dynShape, dynTable);
												if(iTable != null)
												{
													tableFormatHelper.FormatTables(iTable);
													tableCount++;
												}
											}
										}
								}
							}
							catch(System.Exception ex2)
							{
								PPA.Core.Profiler.LogMessage($"dynamic 访问 ShapeRange 也失败: {ex2.Message}", "ERROR");
							}
						}
					}

					if(tableCount == 0)
					{
						if(sel is IShape abstractShapeSel)
						{
							ProcessAbstractShape(abstractShapeSel);
						}
						else if(sel is IEnumerable<IShape> abstractShapesSel)
						{
							foreach(var abstractShape in abstractShapesSel)
							{
								ProcessAbstractShape(abstractShape);
							}
						}
					}


					if(tableCount>0)
					{
						Toast.Show(ResourceManager.GetString("Toast_FormatTables_Success","成功美化 {0} 个表格",tableCount),Toast.ToastType.Success);
						return; // 已处理选中对象，直接返回，不再处理未选中对象
					}

					Toast.Show(ResourceManager.GetString("Toast_FormatTables_NoSelection","选中的对象中没有表格"),Toast.ToastType.Info);
					return; // 选中对象中没有表格，直接返回，不再处理未选中对象
				} else
				{
					// 未选中对象的情况，美化当前幻灯片所有表格
					if(slide!=null)
					{
						PPA.Core.Profiler.LogMessage($"处理幻灯片所有形状，slide类型={slide.GetType().Name}", "INFO");
						try
						{
							// 尝试使用 NetOffice 枚举
							foreach(NETOP.Shape shape in slide.Shapes)
							{
								// 检查是否是表格：先尝试 HasTable，如果失败则直接检查 Table 属性
								bool isTable = false;
								dynamic dynTable = null;
								try
								{
									if(shape.HasTable == MsoTriState.msoTrue)
									{
										isTable = true;
										dynTable = shape.Table;
									}
								}
								catch
								{
									// HasTable 不可用，尝试直接检查 Table 属性
									try
									{
										dynamic dynShape = shape;
										dynTable = SafeGet(() => dynShape.Table, null);
										isTable = (dynTable != null);
									}
									catch { }
								}
								
								if(isTable && dynTable != null)
								{
									PPA.Core.Profiler.LogMessage($"发现表格形状: {shape.Name}", "INFO");
									var iTable = AdapterUtils.WrapTable(app, shape, dynTable);
									PPA.Core.Profiler.LogMessage($"WrapTable 返回: {iTable?.GetType().Name ?? "null"}", "INFO");
									if(iTable != null)
									{
										tableFormatHelper.FormatTables(iTable);
										tableCount++;
									}
									else
									{
										PPA.Core.Profiler.LogMessage("WrapTable 返回 null，无法格式化", "ERROR");
									}
								}
							}
						}
						catch(System.Exception ex)
						{
							// NetOffice 无法枚举 WPS Shapes，使用 dynamic 访问
							PPA.Core.Profiler.LogMessage($"NetOffice 枚举失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
							try
							{
								dynamic dynSlide = slide;
								if(dynSlide == null)
								{
									if(slide is IComWrapper wrapper)
									{
										dynSlide = wrapper.NativeObject;
									}
								}
								
								dynamic dynShapes = SafeGet(() => dynSlide.Shapes, null);
								if(dynShapes != null)
								{
									int count = SafeGet(() => (int)dynShapes.Count, 0);
									PPA.Core.Profiler.LogMessage($"使用 dynamic 访问，Shapes Count={count}", "INFO");
									for(int i = 1; i <= count; i++)
									{
										try
										{
											dynamic dynShape = SafeGet(() => dynShapes[i], null);
											if(dynShape != null)
											{
												// WPS 中 HasTable 可能不可用，直接检查 Table 属性是否存在
												dynamic dynTable = null;
												bool hasTable = false;
												try
												{
													// 方式1：尝试通过 HasTable 属性
													hasTable = SafeGet(() => (bool)(dynShape.HasTable ?? false), false);
													if(hasTable)
													{
														dynTable = SafeGet(() => dynShape.Table, null);
													}
												}
												catch { }
												
												// 方式2：如果方式1失败，直接检查 Table 属性
												if(!hasTable || dynTable == null)
												{
													try
													{
														dynTable = SafeGet(() => dynShape.Table, null);
														hasTable = (dynTable != null);
													}
													catch { }
												}
												
												if(hasTable && dynTable != null)
												{
													PPA.Core.Profiler.LogMessage($"发现表格形状 {i}，Table 属性存在", "INFO");
													var iTable = AdapterUtils.WrapTable(app, dynShape, dynTable);
													if(iTable != null)
													{
														tableFormatHelper.FormatTables(iTable);
														tableCount++;
													}
													else
													{
														PPA.Core.Profiler.LogMessage($"形状 {i} WrapTable 返回 null", "WARN");
													}
												}
											}
										}
										catch(System.Exception ex3)
										{
											PPA.Core.Profiler.LogMessage($"处理形状 {i} 时出错: {ex3.Message}", "WARN");
										}
									}
								}
							}
							catch(System.Exception ex2)
							{
								PPA.Core.Profiler.LogMessage($"dynamic 访问也失败: {ex2.Message}", "ERROR");
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

		private async Task FormatTablesInternalAsync(
			NETOP.Application app,
			IProgress<AsyncProgress> progress,
			CancellationToken cancellationToken,
			ITableFormatHelper tableFormatHelper)
		{
			if(tableFormatHelper==null)
				throw new InvalidOperationException("无法获取 ITableFormatHelper 服务");

			try
			{
				await UndoHelper.BeginUndoEntryAsync(app,UndoHelper.UndoNames.FormatTables);

				var slide = await AsyncOperationHelper.RunOnUIThread(() =>
				{
					return _shapeHelper.TryGetCurrentSlide(app);
				});

				if(slide==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				var tables = await AsyncOperationHelper.RunOnUIThread(() =>
				{
					var sel = _shapeHelper.ValidateSelection(app);
					var result = new List<(NETOP.Shape shape, NETOP.Table table)>();

					if(sel!=null)
					{
						if(sel is NETOP.Shape shape)
						{
							// 检查是否是表格：先尝试 HasTable，如果失败则直接检查 Table 属性
							bool isTable = false;
							dynamic dynTable = null;
							try
							{
								if(shape.HasTable == MsoTriState.msoTrue)
								{
									isTable = true;
									dynTable = shape.Table;
								}
							}
							catch
							{
								// HasTable 不可用，尝试直接检查 Table 属性
								try
								{
									dynamic dynShape = shape;
									dynTable = SafeGet(() => dynShape.Table, null);
									isTable = (dynTable != null);
								}
								catch { }
							}
							
							if(isTable && dynTable != null)
							{
								result.Add((shape, (NETOP.Table)(object)dynTable));
								progress?.Report(new AsyncProgress(10,ResourceManager.GetString("Progress_TableFound","发现表格"),1,1));
							}
						}
						else if(sel is NETOP.ShapeRange shapes)
						{
							try
							{
								int count = 0;
								foreach(NETOP.Shape s in shapes)
								{
									if(s.HasTable==MsoTriState.msoTrue)
									{
										result.Add((s,s.Table));
										count++;
										progress?.Report(new AsyncProgress(
											10,
											ResourceManager.GetString("Progress_TableFound_Count","发现表格 {0}",count),
											count,
											shapes.Count));
									}
								}
							}
							catch(System.Exception ex)
							{
								// NetOffice 无法枚举 WPS ShapeRange，使用 dynamic 访问
								PPA.Core.Profiler.LogMessage($"NetOffice 枚举 ShapeRange 失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
								try
								{
									dynamic dynShapeRange = shapes;
									int rangeCount = SafeGet(() => (int)dynShapeRange.Count, 0);
									int count = 0;
									for(int i = 1; i <= rangeCount; i++)
									{
										dynamic dynShape = SafeGet(() => dynShapeRange[i], null);
										if(dynShape != null)
										{
											// WPS 中 HasTable 可能不可用，直接检查 Table 属性
											dynamic dynTable2 = null;
											bool hasTable = false;
											try
											{
												// 方式1：尝试通过 HasTable 属性
												hasTable = SafeGet(() => (bool)(dynShape.HasTable ?? false), false);
												if(hasTable)
												{
													dynTable2 = SafeGet(() => dynShape.Table, null);
												}
											}
											catch { }
											
											// 方式2：如果方式1失败，直接检查 Table 属性
											if(!hasTable || dynTable2 == null)
											{
												try
												{
													dynTable2 = SafeGet(() => dynShape.Table, null);
													hasTable = (dynTable2 != null);
												}
												catch { }
											}
											
											if(hasTable && dynTable2 != null)
											{
												result.Add(((NETOP.Shape)(object)dynShape, (NETOP.Table)(object)dynTable2));
												count++;
												progress?.Report(new AsyncProgress(
													10,
													ResourceManager.GetString("Progress_TableFound_Count","发现表格 {0}",count),
													count,
													rangeCount));
											}
										}
									}
								}
								catch { }
							}
						}
					}
					else
					{
						try
						{
							foreach(NETOP.Shape shape2 in slide.Shapes)
							{
								if(shape2.HasTable==MsoTriState.msoTrue)
								{
									result.Add((shape2,shape2.Table));
								}
							}
						}
						catch(System.Exception ex)
						{
							// NetOffice 无法枚举 WPS Shapes，使用 dynamic 访问
							PPA.Core.Profiler.LogMessage($"NetOffice 枚举 Shapes 失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
							try
							{
								dynamic dynSlide = slide;
								if(dynSlide == null)
								{
									if(slide is IComWrapper wrapper)
									{
										dynSlide = wrapper.NativeObject;
									}
								}
								
								dynamic dynShapes = SafeGet(() => dynSlide.Shapes, null);
								if(dynShapes != null)
								{
									int count = SafeGet(() => (int)dynShapes.Count, 0);
									for(int i = 1; i <= count; i++)
									{
										dynamic dynShape = SafeGet(() => dynShapes[i], null);
										if(dynShape != null)
										{
											// WPS 中 HasTable 可能不可用，直接检查 Table 属性
											dynamic dynTable3 = null;
											bool hasTable = false;
											try
											{
												// 方式1：尝试通过 HasTable 属性
												hasTable = SafeGet(() => (bool)(dynShape.HasTable ?? false), false);
												if(hasTable)
												{
													dynTable3 = SafeGet(() => dynShape.Table, null);
												}
											}
											catch { }
											
											// 方式2：如果方式1失败，直接检查 Table 属性
											if(!hasTable || dynTable3 == null)
											{
												try
												{
													dynTable3 = SafeGet(() => dynShape.Table, null);
													hasTable = (dynTable3 != null);
												}
												catch { }
											}
											
											if(hasTable && dynTable3 != null)
											{
												result.Add(((NETOP.Shape)(object)dynShape, (NETOP.Table)(object)dynTable3));
											}
										}
									}
								}
							}
							catch { }
						}
					}

					return result;
				});

				cancellationToken.ThrowIfCancellationRequested();

				int total = tables.Count;

				if(total==0)
				{
					var hasSelection = await AsyncOperationHelper.RunOnUIThread(() =>
					{
						return _shapeHelper.ValidateSelection(app)!=null;
					});
					Toast.Show(hasSelection ? ResourceManager.GetString("Toast_FormatTables_NoSelection","选中的对象中没有表格") : ResourceManager.GetString("Toast_FormatTables_NoTables","当前幻灯片上没有表格"),Toast.ToastType.Info);
					return;
				}

				progress?.Report(new AsyncProgress(10,ResourceManager.GetString("Progress_TablesFound","发现 {0} 个表格",total),total,total));
				progress?.Report(new AsyncProgress(20,ResourceManager.GetString("Progress_FormatTables_Start","开始美化 {0} 个表格",total),0,total));

				for(int i = 0;i<total;i++)
				{
					cancellationToken.ThrowIfCancellationRequested();

					NETOP.Shape shape = tables[i].shape;
					NETOP.Table table = tables[i].table;

					int currentProgress = 20 + (int)((i * 60.0) / total);
					progress?.Report(new AsyncProgress(
						currentProgress,
						ResourceManager.GetString("Progress_FormatTable_Progress","美化表格 {0}/{1}",i+1,total),
						i+1,
						total));

					await AsyncOperationHelper.RunOnUIThread(() =>
					{
						var iTable = AdapterUtils.WrapTable(app,shape,table);
						tableFormatHelper.FormatTables(iTable);
					});

					await Task.Delay(10,cancellationToken);
				}

				progress?.Report(new AsyncProgress(100,ResourceManager.GetString("Progress_FormatTables_Complete","成功美化 {0} 个表格",total),total,total));
			} catch
			{
				throw;
			}
		}

		#endregion

		// 适配包装逻辑已提取到 Core/Adapters/AdapterUtils.cs

		private static T SafeGet<T>(System.Func<T> getter, T @default = default)
		{
			try { return getter(); } catch { return @default; }
		}

	}
}

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

		public void FormatTables(NETOP.Application netApp)
		{
			if(netApp==null) throw new ArgumentNullException(nameof(netApp));
			FormatTablesInternal(netApp,_tableFormatHelper);
		}

		public Task FormatTablesAsync(
			NETOP.Application netApp,
			IProgress<AsyncProgress> progress = null,
			CancellationToken cancellationToken = default)
		{
			if(netApp==null) throw new ArgumentNullException(nameof(netApp));
			return FormatTablesInternalAsync(netApp,progress,cancellationToken,_tableFormatHelper);
		}

		#endregion

		#region 内部实现

		private void FormatTablesInternal(NETOP.Application netApp,ITableFormatHelper tableFormatHelper)
		{
			Profiler.LogMessage($"FormatTablesInternal 开始，netApp类型={netApp?.GetType().Name ?? "null"}", "INFO");
			if(tableFormatHelper==null)
				throw new InvalidOperationException("无法获取 ITableFormatHelper 服务");

			UndoHelper.BeginUndoEntry(netApp,UndoHelper.UndoNames.FormatTables);

			var abstractApp = ApplicationHelper.GetAbstractApplication(netApp);
			var slide = _shapeHelper.TryGetCurrentSlide(abstractApp);
			Profiler.LogMessage($"TryGetCurrentSlide 返回: {slide?.GetType().Name ?? "null"}", "INFO");

			ExHandler.Run(() =>
			{
				var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
				Profiler.LogMessage($"ValidateSelection 返回: {sel?.GetType().Name ?? "null"}", "INFO");
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
					Profiler.LogMessage($"处理抽象形状: {shapeName ?? "未知"}, HasTable={abstractShape.HasTable}", "INFO");

					if(!abstractShape.HasTable)
					{
						Profiler.LogMessage($"抽象形状 {shapeName ?? "未知"} 不包含表格，跳过", "INFO");
						return false;
					}

					var key = (abstractShape as IComWrapper)?.NativeObject ?? abstractShape;
					if(AlreadyProcessed(key))
					{
						Profiler.LogMessage($"抽象形状 {shapeName ?? "未知"} 已处理，跳过", "INFO");
						return false;
					}

					var table = abstractShape.GetTable();
					if(table != null)
					{
						Profiler.LogMessage($"抽象形状 {shapeName ?? "未知"} 返回表格实例: {table.GetType().Name}", "INFO");
						tableFormatHelper.FormatTables(table);
						tableCount++;
						return true;
					}

					Profiler.LogMessage($"抽象形状 {shapeName ?? "未知"} GetTable 返回 null", "WARN");
					return false;
				}


				if(sel!=null)
				{
					// 有选中对象的情况，美化选中对象中的表格
					if(sel is NETOP.Shape shape)
					{
						if(AlreadyProcessed(shape))
						{
							Profiler.LogMessage($"形状 {shape.Name} 已经处理，跳过", "INFO");
						}
						else
						{
						// 检查是否是表格：先尝试 HasTable，如果失败则直接检查 Table 属性
						bool isTable = false;
						dynamic dynTable = null;
						try
						{
								Profiler.LogMessage($"选中单个形状，HasTable={shape.HasTable}", "INFO");
							if(shape.HasTable == MsoTriState.msoTrue)
							{
								isTable = true;
								dynTable = shape.Table;
							}
						}
						catch
						{
								// HasTable 不可用，尝试直接检查 Table 属性
								Profiler.LogMessage("HasTable 不可用，尝试直接检查 Table 属性", "INFO");
							try
							{
								dynamic dynShape = shape;
								dynTable = SafeGet(() => dynShape.Table, null);
								isTable = (dynTable != null);
									Profiler.LogMessage($"直接检查 Table 属性: isTable={isTable}", "INFO");
							}
							catch(System.Exception ex)
							{
									Profiler.LogMessage($"检查 Table 属性失败: {ex.Message}", "WARN");
							}
						}
						
						if(isTable && dynTable != null)
						{
								Profiler.LogMessage("开始包装表格", "INFO");
							var iTable = AdapterUtils.WrapTable(netApp, shape, dynTable);
								Profiler.LogMessage($"WrapTable 返回: {iTable?.GetType().Name ?? "null"}", "INFO");
							if(iTable != null)
							{
								tableFormatHelper.FormatTables(iTable);
								tableCount++;
							}
							else
							{
									Profiler.LogMessage("WrapTable 返回 null，无法格式化", "ERROR");
							}
						}
						else
						{
								Profiler.LogMessage("形状不是表格或 Table 属性为 null", "INFO");
						}
						}
					} else if(sel is NETOP.ShapeRange shapes)
					{
						Profiler.LogMessage($"选中多个形状，Count={shapes.Count}", "INFO");
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
										Profiler.LogMessage($"形状 {s.Name} 已处理，跳过", "INFO");
										continue;
									}
									Profiler.LogMessage($"处理形状 {s.Name}，检测到表格", "INFO");
									var iTable = AdapterUtils.WrapTable(netApp, s, dynTable);
									Profiler.LogMessage($"WrapTable 返回: {iTable?.GetType().Name ?? "null"}", "INFO");
									if(iTable != null)
									{
										tableFormatHelper.FormatTables(iTable);
										tableCount++;
									}
									else
									{
										Profiler.LogMessage("WrapTable 返回 null，无法格式化", "ERROR");
									}
								}
							}
						}
						catch(System.Exception ex)
						{
							// NetOffice 无法枚举某些 ShapeRange，使用 dynamic 访问作为后备方案
							Profiler.LogMessage($"NetOffice 枚举 ShapeRange 失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
							try
							{
								dynamic dynShapeRange = shapes;
								int rangeCount = SafeGet(() => (int)dynShapeRange.Count, 0);
								Profiler.LogMessage($"使用 dynamic 访问 ShapeRange，Count={rangeCount}", "INFO");
								for(int i = 1; i <= rangeCount; i++)
								{
										dynamic dynShape = SafeGet(() => dynShapeRange[i], null);
										if(dynShape != null)
										{
											// 某些情况下 HasTable 可能不可用，直接检查 Table 属性
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
												Profiler.LogMessage($"dynamic 形状 {i} 已处理，跳过", "INFO");
												continue;
											}
											Profiler.LogMessage($"发现表格形状 {i}", "INFO");
												var iTable = AdapterUtils.WrapTable(netApp, dynShape, dynTable);
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
								Profiler.LogMessage($"dynamic 访问 ShapeRange 也失败: {ex2.Message}", "ERROR");
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
						Profiler.LogMessage($"处理幻灯片所有形状，slide类型={slide.GetType().Name}", "INFO");
						try
						{
							// 从 ISlide 获取底层的 NETOP.Slide 对象
							NETOP.Slide nativeSlide = null;
							if(slide is IComWrapper<NETOP.Slide> typedSlide)
							{
								nativeSlide = typedSlide.NativeObject;
							}
							else if(slide is IComWrapper wrapper)
							{
								nativeSlide = wrapper.NativeObject as NETOP.Slide;
							}

							if(nativeSlide == null)
							{
								Profiler.LogMessage("无法获取底层 NETOP.Slide 对象", "WARN");
								throw new InvalidOperationException("无法获取底层 NETOP.Slide 对象");
							}

							// 尝试使用 NetOffice 枚举
							foreach(NETOP.Shape shape in nativeSlide.Shapes)
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
									Profiler.LogMessage($"发现表格形状: {shape.Name}", "INFO");
									var iTable = AdapterUtils.WrapTable(netApp, shape, dynTable);
									Profiler.LogMessage($"WrapTable 返回: {iTable?.GetType().Name ?? "null"}", "INFO");
									if(iTable != null)
									{
										tableFormatHelper.FormatTables(iTable);
										tableCount++;
									}
									else
									{
										Profiler.LogMessage("WrapTable 返回 null，无法格式化", "ERROR");
									}
								}
							}
						}
						catch(System.Exception ex)
						{
							// NetOffice 无法枚举某些 Shapes，使用 dynamic 访问作为后备方案
							Profiler.LogMessage($"NetOffice 枚举失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
							try
							{
								// 从 ISlide 获取底层的 NETOP.Slide 对象
								NETOP.Slide nativeSlide = null;
								if(slide is IComWrapper<NETOP.Slide> typedSlide)
								{
									nativeSlide = typedSlide.NativeObject;
								}
								else if(slide is IComWrapper wrapper)
								{
									nativeSlide = wrapper.NativeObject as NETOP.Slide;
								}

								if(nativeSlide == null)
								{
									Profiler.LogMessage("无法获取底层 NETOP.Slide 对象（dynamic 访问）", "WARN");
									throw new InvalidOperationException("无法获取底层 NETOP.Slide 对象");
								}

								dynamic dynSlide = nativeSlide;
								dynamic dynShapes = SafeGet(() => dynSlide.Shapes, null);
								if(dynShapes != null)
								{
									int count = SafeGet(() => (int)dynShapes.Count, 0);
									Profiler.LogMessage($"使用 dynamic 访问，Shapes Count={count}", "INFO");
									for(int i = 1; i <= count; i++)
									{
										try
										{
											object shapeObj = SafeGet(() => dynShapes[i], null);
											if(shapeObj == null) continue;

											// 尝试转换为 NETOP.Shape
											NETOP.Shape netShape = null;
											if(shapeObj is NETOP.Shape directShape)
											{
												netShape = directShape;
											}
											else if(shapeObj is IComWrapper<NETOP.Shape> typedShape)
											{
												netShape = typedShape.NativeObject;
											}
											else if(shapeObj is IComWrapper wrapper)
											{
												netShape = wrapper.NativeObject as NETOP.Shape;
											}
											else
											{
												// 尝试强制转换
												try
												{
													netShape = (NETOP.Shape)(object)shapeObj;
												}
												catch { }
											}

											if(netShape == null) continue;

											// 检查是否是表格
											bool isTable = false;
											dynamic dynTable = null;
											try
											{
												if(netShape.HasTable == MsoTriState.msoTrue)
												{
													isTable = true;
													dynTable = netShape.Table;
												}
											}
											catch
											{
												// HasTable 不可用，尝试直接检查 Table 属性
												try
												{
													dynTable = netShape.Table;
													isTable = (dynTable != null);
												}
												catch { }
											}

											if(isTable && dynTable != null)
											{
												Profiler.LogMessage($"发现表格形状 {i}: {netShape.Name}", "INFO");
												var iTable = AdapterUtils.WrapTable(netApp, netShape, dynTable);
												if(iTable != null)
												{
													tableFormatHelper.FormatTables(iTable);
													tableCount++;
												}
												else
												{
													Profiler.LogMessage($"形状 {i} WrapTable 返回 null", "WARN");
												}
											}
										}
										catch(System.Exception ex3)
										{
											Profiler.LogMessage($"处理形状 {i} 时出错: {ex3.Message}", "WARN");
										}
									}
								}
							}
							catch(System.Exception ex2)
							{
								Profiler.LogMessage($"dynamic 访问也失败: {ex2.Message}", "ERROR");
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
			NETOP.Application netApp,
			IProgress<AsyncProgress> progress,
			CancellationToken cancellationToken,
			ITableFormatHelper tableFormatHelper)
		{
			if(tableFormatHelper==null)
				throw new InvalidOperationException("无法获取 ITableFormatHelper 服务");

			try
			{
				await UndoHelper.BeginUndoEntryAsync(netApp,UndoHelper.UndoNames.FormatTables);

				var abstractAppForSlide = ApplicationHelper.GetAbstractApplication(netApp);
				var slide = await AsyncOperationHelper.RunOnUIThread(() =>
				{
					return _shapeHelper.TryGetCurrentSlide(abstractAppForSlide);
				});

				if(slide==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				var tables = await AsyncOperationHelper.RunOnUIThread(() =>
				{
					var sel = _shapeHelper.ValidateSelection(abstractAppForSlide) as dynamic;
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
								// NetOffice 无法枚举某些 ShapeRange，使用 dynamic 访问作为后备方案
								Profiler.LogMessage($"NetOffice 枚举 ShapeRange 失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
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
											// 某些情况下 HasTable 可能不可用，直接检查 Table 属性
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
							// 从 ISlide 获取底层的 NETOP.Slide 对象
							NETOP.Slide nativeSlide = null;
							if(slide is IComWrapper<NETOP.Slide> typedSlide)
							{
								nativeSlide = typedSlide.NativeObject;
							}
							else if(slide is IComWrapper wrapper)
							{
								nativeSlide = wrapper.NativeObject as NETOP.Slide;
							}

							if(nativeSlide == null)
							{
								Profiler.LogMessage("无法获取底层 NETOP.Slide 对象（异步方法）", "WARN");
								throw new InvalidOperationException("无法获取底层 NETOP.Slide 对象");
							}

							foreach(NETOP.Shape shape2 in nativeSlide.Shapes)
							{
								if(shape2.HasTable==MsoTriState.msoTrue)
								{
									result.Add((shape2,shape2.Table));
								}
							}
						}
						catch(System.Exception ex)
						{
							// NetOffice 无法枚举某些 Shapes，使用 dynamic 访问作为后备方案
							Profiler.LogMessage($"NetOffice 枚举 Shapes 失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
							try
							{
								// 从 ISlide 获取底层的 NETOP.Slide 对象
								NETOP.Slide nativeSlide = null;
								if(slide is IComWrapper<NETOP.Slide> typedSlide)
								{
									nativeSlide = typedSlide.NativeObject;
								}
								else if(slide is IComWrapper wrapper)
								{
									nativeSlide = wrapper.NativeObject as NETOP.Slide;
								}

								if(nativeSlide == null)
								{
									Profiler.LogMessage("无法获取底层 NETOP.Slide 对象（异步方法 dynamic 访问）", "WARN");
									throw new InvalidOperationException("无法获取底层 NETOP.Slide 对象");
								}

								dynamic dynSlide = nativeSlide;
								dynamic dynShapes = SafeGet(() => dynSlide.Shapes, null);
								if(dynShapes != null)
								{
									int count = SafeGet(() => (int)dynShapes.Count, 0);
									for(int i = 1; i <= count; i++)
									{
										object shapeObj = SafeGet(() => dynShapes[i], null);
										if(shapeObj == null) continue;

										// 尝试转换为 NETOP.Shape
										NETOP.Shape netShape = null;
										if(shapeObj is NETOP.Shape directShape)
										{
											netShape = directShape;
										}
										else if(shapeObj is IComWrapper<NETOP.Shape> typedShape)
										{
											netShape = typedShape.NativeObject;
										}
										else if(shapeObj is IComWrapper wrapper)
										{
											netShape = wrapper.NativeObject as NETOP.Shape;
										}
										else
										{
											try
											{
												netShape = (NETOP.Shape)(object)shapeObj;
											}
											catch { }
										}

										if(netShape == null) continue;

										// 检查是否是表格
										bool hasTable = false;
										NETOP.Table netTable = null;
										try
										{
											if(netShape.HasTable == MsoTriState.msoTrue)
											{
												hasTable = true;
												netTable = netShape.Table;
											}
										}
										catch
										{
											try
											{
												netTable = netShape.Table;
												hasTable = (netTable != null);
											}
											catch { }
										}

										if(hasTable && netTable != null)
										{
											result.Add((netShape, netTable));
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
					var abstractAppForCheck = ApplicationHelper.GetAbstractApplication(netApp);
					var hasSelection = await AsyncOperationHelper.RunOnUIThread(() =>
					{
						return _shapeHelper.ValidateSelection(abstractAppForCheck)!=null;
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
						var iTable = AdapterUtils.WrapTable(netApp,shape,table);
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

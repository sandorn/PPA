using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using PPA.Shape;
using PPA.Utilities;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 图表批量操作辅助类
	/// </summary>
	internal class ChartBatchHelper : IChartBatchHelper
	{
		private readonly IChartFormatHelper _chartFormatHelper;
		private readonly IShapeHelper _shapeHelper;

		public ChartBatchHelper(IChartFormatHelper chartFormatHelper, IShapeHelper shapeHelper)
		{
			_chartFormatHelper = chartFormatHelper ?? throw new ArgumentNullException(nameof(chartFormatHelper));
			_shapeHelper = shapeHelper ?? throw new ArgumentNullException(nameof(shapeHelper));
		}

		#region IChartBatchHelper 实现

		public void FormatCharts(NETOP.Application app)
		{
			if(app==null) throw new ArgumentNullException(nameof(app));
			FormatChartsInternal(app,_chartFormatHelper);
		}

		public Task FormatChartsAsync(NETOP.Application app, IProgress<AsyncProgress> progress = null)
		{
			FormatCharts(app);
			progress?.Report(new AsyncProgress(100,ResourceManager.GetString("Progress_FormatCharts_Complete","图表美化完成"),1,1));
			return Task.CompletedTask;
		}

		#endregion

		#region 向后兼容的静态入口

		public static void Bt503_Click(NETOP.Application app, IChartFormatHelper chartFormatHelper = null)
		{
			var helper = ResolveInstance(chartFormatHelper);
			helper.FormatCharts(app);
		}

		private static IChartBatchHelper ResolveInstance(IChartFormatHelper overrideHelper)
		{
			if(overrideHelper!=null)
			{
				return new ChartBatchHelper(overrideHelper,ResolveShapeHelper());
			}

			var serviceProvider = Globals.ThisAddIn?.ServiceProvider;
			if(serviceProvider!=null)
			{
				var resolved = serviceProvider.GetService<IChartBatchHelper>();
				if(resolved!=null) return resolved;

				var formatHelper = serviceProvider.GetService<IChartFormatHelper>();
				var shapeHelper = serviceProvider.GetService<IShapeHelper>();
				if(formatHelper!=null&&shapeHelper!=null)
				{
					return new ChartBatchHelper(formatHelper,shapeHelper);
				}
			}

			return new ChartBatchHelper(
				new ChartFormatHelper(FormattingConfig.Instance,ResolveShapeHelper()),
				ResolveShapeHelper());
		}

		private static IShapeHelper ResolveShapeHelper()
		{
			var serviceProvider = Globals.ThisAddIn?.ServiceProvider;
			var helper = serviceProvider?.GetService<IShapeHelper>();
			return helper ?? ShapeUtils.Default;
		}

		private static T SafeGet<T>(System.Func<T> getter, T @default = default)
		{
			try { return getter(); } catch { return @default; }
		}

		#endregion

		#region 内部实现

		private void FormatChartsInternal(NETOP.Application app, IChartFormatHelper chartFormatHelper)
		{
			PPA.Core.Profiler.LogMessage($"FormatChartsInternal 开始，app类型={app?.GetType().Name ?? "null"}", "INFO");
			if(chartFormatHelper==null)
				throw new InvalidOperationException("无法获取 IChartFormatHelper 服务");

			UndoHelper.BeginUndoEntry(app,UndoHelper.UndoNames.FormatCharts);

			ExHandler.Run(() =>
			{
				var slide = _shapeHelper.TryGetCurrentSlide(app);
				PPA.Core.Profiler.LogMessage($"TryGetCurrentSlide 返回: {slide?.GetType().Name ?? "null"}", "INFO");
				if(slide==null) return;

				var chartShapes = new List<NETOP.Shape>();
				var selection = _shapeHelper.ValidateSelection(app);
				PPA.Core.Profiler.LogMessage($"ValidateSelection 返回: {selection?.GetType().Name ?? "null"}", "INFO");

				// 动态选区兜底：当 ValidateSelection 返回 null 时，直接从 ActiveWindow.Selection 读取
				if(selection == null)
				{
					try
					{
						dynamic dynApp = app;
						dynamic activeWindow = SafeGet(() => dynApp.ActiveWindow, null);
						if(activeWindow != null)
						{
							dynamic dynSelection = SafeGet(() => activeWindow.Selection, null);
							if(dynSelection != null)
							{
								// 尝试获取 ShapeRange
								dynamic shapeRange = SafeGet(() => dynSelection.ShapeRange, null);
								if(shapeRange != null)
								{
									int rangeCount = SafeGet(() => (int)shapeRange.Count, 0);
									if(rangeCount > 0)
									{
										PPA.Core.Profiler.LogMessage($"动态选区兜底：从 ActiveWindow.Selection.ShapeRange 获取到 {rangeCount} 个形状", "INFO");
										for(int i = 1; i <= rangeCount; i++)
										{
											dynamic dynShape = SafeGet(() => shapeRange[i], null);
											if(dynShape != null)
											{
												dynamic dynChart = SafeGet(() => dynShape.Chart, null);
												bool hasChart = SafeGet(() => (bool)(dynShape.HasChart ?? false), false) || (dynChart != null);
												if(hasChart && dynChart != null)
												{
													try
													{
														chartShapes.Add((NETOP.Shape)(object)dynShape);
													}
													catch { }
												}
											}
										}
									}
								}
								
								// 如果 ShapeRange 为空，尝试获取单个 Shape
								if(chartShapes.Count == 0)
								{
									dynamic singleShape = SafeGet(() => dynSelection.Shape, null);
									if(singleShape != null)
									{
										dynamic dynChart = SafeGet(() => singleShape.Chart, null);
										bool hasChart = SafeGet(() => (bool)(singleShape.HasChart ?? false), false) || (dynChart != null);
										if(hasChart && dynChart != null)
										{
											PPA.Core.Profiler.LogMessage("动态选区兜底：从 ActiveWindow.Selection.Shape 获取到单个图表形状", "INFO");
											try
											{
												chartShapes.Add((NETOP.Shape)(object)singleShape);
											}
											catch { }
										}
									}
								}
							}
						}
					}
					catch(System.Exception ex)
					{
						PPA.Core.Profiler.LogMessage($"动态选区兜底失败: {ex.Message}", "WARN");
					}
				}


				if(selection!=null)
				{
					if(selection is NETOP.Shape shape)
					{
						// 检查是否是图表：先尝试 HasChart，如果失败或返回 false，则直接检查 Chart 属性
						bool isChart = false;
						dynamic dynChart = null;
						try
						{
							if(shape.HasChart == MsoTriState.msoTrue)
							{
								isChart = true;
								dynChart = shape.Chart;
							}
						}
						catch { }
						
						// 如果 HasChart 不可用或返回 false，直接检查 Chart 属性
						if(!isChart || dynChart == null)
						{
							try
							{
								dynamic dynShape = shape;
								dynChart = SafeGet(() => dynShape.Chart, null);
								isChart = (dynChart != null);
								PPA.Core.Profiler.LogMessage($"单选形状，HasChart 检查失败，直接检查 Chart 属性: isChart={isChart}", "INFO");
							}
							catch(System.Exception ex)
							{
								PPA.Core.Profiler.LogMessage($"检查 Chart 属性失败: {ex.Message}", "WARN");
							}
						}
						
						if(isChart && dynChart != null)
						{
							PPA.Core.Profiler.LogMessage($"单选形状检测到图表，添加到列表", "INFO");
							chartShapes.Add(shape);
						}
						else
						{
							PPA.Core.Profiler.LogMessage($"单选形状不是图表: isChart={isChart}, dynChart={dynChart != null}", "INFO");
						}
					}
					else if(selection is NETOP.ShapeRange shapeRange)
					{
						try
						{
							foreach(NETOP.Shape s in shapeRange)
							{
								// 检查是否是图表：先尝试 HasChart，如果失败则直接检查 Chart 属性
								bool isChart = false;
								try
								{
									if(s.HasChart == MsoTriState.msoTrue)
									{
										isChart = true;
									}
								}
								catch
								{
									// HasChart 不可用，尝试直接检查 Chart 属性
									try
									{
										dynamic dynShape = s;
										var dynChart = SafeGet(() => dynShape.Chart, null);
										isChart = (dynChart != null);
									}
									catch { }
								}
								
								if(isChart)
								{
									chartShapes.Add(s);
								}
							}
						}
						catch(System.Exception ex)
						{
							// NetOffice 无法枚举 WPS ShapeRange，使用 dynamic 访问
							PPA.Core.Profiler.LogMessage($"NetOffice 枚举 ShapeRange 失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
							try
							{
								dynamic dynShapeRange = shapeRange;
								int rangeCount = SafeGet(() => (int)dynShapeRange.Count, 0);
								PPA.Core.Profiler.LogMessage($"使用 dynamic 访问 ShapeRange，Count={rangeCount}", "INFO");
								for(int i = 1; i <= rangeCount; i++)
								{
									dynamic dynShape = SafeGet(() => dynShapeRange[i], null);
									if(dynShape != null)
									{
										// WPS 中 HasChart 可能不可用，直接检查 Chart 属性是否存在
										dynamic dynChart = null;
										bool hasChart = false;
										try
										{
											// 方式1：尝试通过 HasChart 属性
											hasChart = SafeGet(() => (bool)(dynShape.HasChart ?? false), false);
											if(hasChart)
											{
												dynChart = SafeGet(() => dynShape.Chart, null);
											}
										}
										catch { }
										
										// 方式2：如果方式1失败，直接检查 Chart 属性
										if(!hasChart || dynChart == null)
										{
											try
											{
												dynChart = SafeGet(() => dynShape.Chart, null);
												hasChart = (dynChart != null);
											}
											catch { }
										}
										
										if(hasChart && dynChart != null)
										{
											try
											{
												chartShapes.Add((NETOP.Shape)(object)dynShape);
											}
											catch
											{
												// 转换失败，尝试通过 AdapterUtils 包装
												try
												{
													var iShape = AdapterUtils.WrapShape(app, dynShape);
													if(iShape != null && iShape.HasChart)
													{
														if(iShape is IComWrapper wrapper)
														{
															var nativeShape = wrapper.NativeObject;
															if(nativeShape is NETOP.Shape netShape2)
															{
																chartShapes.Add(netShape2);
															}
														}
													}
												}
												catch { }
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
					else if(selection is IShape abstractShapeSelection)
					{
						// 尝试从抽象形状获取 PowerPoint Shape
						if(abstractShapeSelection is IComWrapper<NETOP.Shape> typedShape)
						{
							var pptShape = typedShape.NativeObject;
							if(pptShape.HasChart == MsoTriState.msoTrue)
							{
								chartShapes.Add(pptShape);
							}
						}
					}
					else if(selection is IEnumerable<IShape> abstractShapeEnumerable)
					{
						foreach(var abstractShape in abstractShapeEnumerable)
						{
							// 尝试从抽象形状获取 PowerPoint Shape
							if(abstractShape is IComWrapper<NETOP.Shape> typedShape)
							{
								var pptShape = typedShape.NativeObject;
								if(pptShape.HasChart == MsoTriState.msoTrue)
								{
									chartShapes.Add(pptShape);
								}
							}
						}
					}
				}
				else
				{
					// 如果从动态选区兜底获取到图表，只处理这些图表并立即返回
					if(chartShapes.Count > 0)
					{
						PPA.Core.Profiler.LogMessage($"动态选区兜底获取到 {chartShapes.Count} 个图表，只处理这些图表", "INFO");
					}

					// 如果已有选中的图表，只处理这些，不再处理整页
					if(chartShapes.Count > 0)
					{
						// 跳过整页枚举，直接处理已选中的图表
					}
					else
					{
						try
						{
							// 尝试使用 NetOffice 枚举
							foreach(NETOP.Shape shape in slide.Shapes)
							{
								// 检查是否是图表：先尝试 HasChart，如果失败则直接检查 Chart 属性
								bool isChart = false;
								try
								{
									if(shape.HasChart == MsoTriState.msoTrue)
									{
										isChart = true;
									}
								}
								catch
								{
									// HasChart 不可用，尝试直接检查 Chart 属性
									try
									{
										dynamic dynShape = shape;
										var dynChart = SafeGet(() => dynShape.Chart, null);
										isChart = (dynChart != null);
									}
									catch { }
								}
								
								if(isChart)
								{
									chartShapes.Add(shape);
								}
							}
						}
					catch(System.Exception ex)
					{
						// NetOffice 无法枚举 WPS Shapes，使用 dynamic 访问
						PPA.Core.Profiler.LogMessage($"NetOffice 枚举失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
						try
						{
							// 尝试从 slide 获取底层 COM 对象
							dynamic dynSlide = slide;
							if(dynSlide == null)
							{
								// 如果 slide 是 NETOP.Slide，尝试获取底层对象
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
								int chartCount = 0;
								for(int i = 1; i <= count; i++)
								{
									try
									{
										dynamic dynShape = SafeGet(() => dynShapes[i], null);
										if(dynShape != null)
										{
											// WPS 中 HasChart 可能不可用，需要直接检查 Chart 属性
											dynamic dynChart = null;
											bool hasChart = false;
											try
											{
												// 方式1：尝试通过 HasChart 属性
												hasChart = SafeGet(() => (bool)(dynShape.HasChart ?? false), false);
												if(hasChart)
												{
													dynChart = SafeGet(() => dynShape.Chart, null);
												}
											}
											catch { }
											
											// 方式2：如果方式1失败，直接检查 Chart 属性
											if(!hasChart || dynChart == null)
											{
												try
												{
													dynChart = SafeGet(() => dynShape.Chart, null);
													hasChart = (dynChart != null);
												}
												catch { }
											}
											
											PPA.Core.Profiler.LogMessage($"形状 {i}: HasChart={hasChart}", "INFO");
											if(hasChart && dynChart != null)
											{
												// 尝试将 dynamic shape 转换为 NETOP.Shape
												// 如果转换失败，直接使用 dynamic 对象
												try
												{
													NETOP.Shape netShape = (NETOP.Shape)(object)dynShape;
													chartShapes.Add(netShape);
													chartCount++;
												}
												catch
												{
													// 如果转换失败，尝试通过 AdapterUtils 包装
													try
													{
														var iShape = AdapterUtils.WrapShape(app, dynShape);
														if(iShape != null && iShape.HasChart)
														{
															// 从 IShape 获取底层 NETOP.Shape
															if(iShape is IComWrapper wrapper)
															{
																var nativeShape = wrapper.NativeObject;
																if(nativeShape is NETOP.Shape netShape2)
																{
																	chartShapes.Add(netShape2);
																	chartCount++;
																}
																else
																{
																	PPA.Core.Profiler.LogMessage($"形状 {i} 的 NativeObject 不是 NETOP.Shape", "WARN");
																}
															}
														}
													}
													catch(System.Exception ex4)
													{
														PPA.Core.Profiler.LogMessage($"包装形状 {i} 失败: {ex4.Message}", "WARN");
													}
												}
											}
										}
									}
									catch(System.Exception ex3)
									{
										PPA.Core.Profiler.LogMessage($"处理形状 {i} 时出错: {ex3.Message}", "WARN");
									}
								}
								PPA.Core.Profiler.LogMessage($"dynamic 访问完成，找到 {chartCount} 个图表", "INFO");
							}
							else
							{
								PPA.Core.Profiler.LogMessage("无法获取 Shapes 集合", "ERROR");
							}
						}
						catch(System.Exception ex2)
						{
							PPA.Core.Profiler.LogMessage($"dynamic 访问也失败: {ex2.Message}", "ERROR");
							PPA.Core.Profiler.LogMessage($"堆栈跟踪: {ex2.StackTrace}", "ERROR");
						}
					}
					}
				}

				var totalCharts = chartShapes.Count;
				PPA.Core.Profiler.LogMessage($"找到 {totalCharts} 个图表形状（selection={selection?.GetType().Name ?? "null"}）", "INFO");
				if(totalCharts == 0)
				{
					PPA.Core.Profiler.LogMessage("没有找到图表，直接返回", "INFO");
					if(selection!=null)
					{
						Toast.Show(ResourceManager.GetString("Toast_FormatCharts_NoSelection"),Toast.ToastType.Info);
					}
					else
					{
						Toast.Show(ResourceManager.GetString("Toast_FormatCharts_NoCharts"),Toast.ToastType.Info);
					}
					return; // 没有图表，直接返回
				}
				foreach(var shape in chartShapes)
				{
					PPA.Core.Profiler.LogMessage($"处理图表形状: {shape.Name}", "INFO");
					var iShape = AdapterUtils.WrapShape(app,shape);
					PPA.Core.Profiler.LogMessage($"WrapShape 返回: {iShape?.GetType().Name ?? "null"}", "INFO");
					if(iShape != null)
					{
						chartFormatHelper.FormatChartText(iShape);
					}
					else
					{
						PPA.Core.Profiler.LogMessage("WrapShape 返回 null，无法格式化", "ERROR");
					}
				}


				if(totalCharts>0)
				{
					Toast.Show(ResourceManager.GetString("Toast_FormatCharts_Success",totalCharts),Toast.ToastType.Success);
				}
				else
				{
					var message = selection!=null
						? ResourceManager.GetString("Toast_FormatCharts_NoSelection")
						: ResourceManager.GetString("Toast_FormatCharts_NoCharts");
					Toast.Show(message,Toast.ToastType.Info);
				}
			});
		}


		#endregion

		// 适配包装逻辑已提取到 Core/Adapters/AdapterUtils.cs
	}
}


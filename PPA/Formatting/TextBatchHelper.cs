using System;
using System.Threading.Tasks;
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
	/// 文本批量操作辅助类
	/// </summary>
	internal class TextBatchHelper : ITextBatchHelper
	{
		private readonly ITextFormatHelper _textFormatHelper;
		private readonly IShapeHelper _shapeHelper;

		public TextBatchHelper(ITextFormatHelper textFormatHelper, IShapeHelper shapeHelper)
		{
			_textFormatHelper = textFormatHelper ?? throw new ArgumentNullException(nameof(textFormatHelper));
			_shapeHelper = shapeHelper ?? throw new ArgumentNullException(nameof(shapeHelper));
		}

		#region ITextBatchHelper 实现

		public void FormatText(NETOP.Application app)
		{
			if(app==null) throw new ArgumentNullException(nameof(app));
			FormatTextInternal(app,_textFormatHelper);
		}

		public Task FormatTextAsync(NETOP.Application app, IProgress<AsyncProgress> progress = null)
		{
			// 当前无异步实现，保持同步调用，可在未来扩展
			FormatText(app);
			progress?.Report(new AsyncProgress(100,ResourceManager.GetString("Progress_FormatText_Complete","文本美化完成"),1,1));
			return Task.CompletedTask;
		}

		#endregion

		#region 内部实现

		private static T SafeGet<T>(System.Func<T> getter, T @default = default)
		{
			try { return getter(); } catch { return @default; }
		}

		#endregion

		#region 内部实现

		private void FormatTextInternal(NETOP.Application netApp,ITextFormatHelper textFormatHelper)
		{
			PPA.Core.Profiler.LogMessage($"FormatTextInternal 开始，netApp类型={netApp?.GetType().Name ?? "null"}", "INFO");
			if(textFormatHelper==null)
				throw new System.InvalidOperationException("无法获取 ITextFormatHelper 服务");

			UndoHelper.BeginUndoEntry(netApp,UndoHelper.UndoNames.FormatText);

			ExHandler.Run(() =>
			{
				var abstractApp = ApplicationHelper.GetAbstractApplication(netApp);
				var selection = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
				PPA.Core.Profiler.LogMessage($"ValidateSelection 返回: {selection?.GetType().Name ?? "null"}", "INFO");

				if(selection==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_FormatText_NoSelection"),Toast.ToastType.Warning);
					return;
				}

				bool hasFormatted = false;

				if(selection is NETOP.Shape shape)
				{
					PPA.Core.Profiler.LogMessage($"选中单个形状，HasText={shape.TextFrame?.HasText}, HasTable={shape.HasTable}", "INFO");
					// 跳过表格内的文本（表格有自己的格式化逻辑）
					if(shape.HasTable==MsoTriState.msoTrue)
					{
						PPA.Core.Profiler.LogMessage("形状包含表格，跳过文本格式化", "INFO");
						return;
					}
					if(shape.TextFrame?.HasText==MsoTriState.msoTrue)
					{
						PPA.Core.Profiler.LogMessage("开始包装形状", "INFO");
						var iShape = AdapterUtils.WrapShape(netApp,shape);
						PPA.Core.Profiler.LogMessage($"WrapShape 返回: {iShape?.GetType().Name ?? "null"}", "INFO");
						if(iShape != null)
						{
							textFormatHelper.ApplyTextFormatting(iShape);
							hasFormatted=true;
						}
						else
						{
							PPA.Core.Profiler.LogMessage("WrapShape 返回 null，无法格式化", "ERROR");
						}
					}
				}
				else if(selection is NETOP.ShapeRange shapeRange)
				{
					PPA.Core.Profiler.LogMessage($"选中多个形状，Count={shapeRange.Count}", "INFO");
					try
					{
						// 尝试使用 NetOffice 枚举
						foreach(NETOP.Shape s in shapeRange)
						{
							if(s.TextFrame?.HasText==MsoTriState.msoTrue)
							{
								PPA.Core.Profiler.LogMessage($"处理形状 {s.Name}，HasText=True", "INFO");
								var iShape = AdapterUtils.WrapShape(netApp,s);
								PPA.Core.Profiler.LogMessage($"WrapShape 返回: {iShape?.GetType().Name ?? "null"}", "INFO");
								if(iShape != null)
								{
									textFormatHelper.ApplyTextFormatting(iShape);
									hasFormatted=true;
								}
								else
								{
									PPA.Core.Profiler.LogMessage("WrapShape 返回 null，无法格式化", "ERROR");
								}
							}
						}
					}
					catch(System.Exception ex)
					{
						// NetOffice 无法枚举某些 ShapeRange，使用 dynamic 访问作为后备方案
						PPA.Core.Profiler.LogMessage($"NetOffice 枚举失败: {ex.Message}，尝试使用 dynamic 访问", "WARN");
						try
						{
							dynamic dynShapeRange = shapeRange;
							int count = SafeGet(() => (int)dynShapeRange.Count, 0);
							PPA.Core.Profiler.LogMessage($"使用 dynamic 访问，Count={count}", "INFO");
							int processedCount = 0;
							for(int i = 1; i <= count; i++)
							{
								try
								{
									dynamic dynShape = SafeGet(() => dynShapeRange[i], null);
									if(dynShape != null)
									{
										// 跳过表格内的文本（表格有自己的格式化逻辑）
										// 某些情况下 HasTable 可能不可用，需要直接检查 Table 属性
										bool hasTable = false;
										try
										{
											hasTable = SafeGet(() => (bool)(dynShape.HasTable ?? false), false);
										}
										catch { }
										
										if(!hasTable)
										{
											// 如果 HasTable 不可用，直接检查 Table 属性
											try
											{
												dynamic dynTable = SafeGet(() => dynShape.Table, null);
												hasTable = (dynTable != null);
											}
											catch { }
										}
										
										if(hasTable)
										{
											PPA.Core.Profiler.LogMessage($"形状 {i} 包含表格，跳过文本格式化", "INFO");
											continue;
										}
										
										// 某些情况下 HasText 可能不可用，尝试多种方式检测
										bool hasText = false;
										try
										{
											// 方式1：尝试 HasText 属性
											hasText = SafeGet(() => (bool)(dynShape.TextFrame?.HasText ?? false), false);
										}
										catch { }
										
										if(!hasText)
										{
											// 方式2：检查 TextFrame 是否存在且有文本
											try
											{
												dynamic textFrame = SafeGet(() => dynShape.TextFrame, null);
												if(textFrame != null)
												{
													dynamic textRange = SafeGet(() => textFrame.TextRange, null);
													if(textRange != null)
													{
														string text = SafeGet(() => (string)textRange.Text, null);
														hasText = !string.IsNullOrWhiteSpace(text);
													}
												}
											}
											catch { }
										}
										
										PPA.Core.Profiler.LogMessage($"形状 {i}: HasText={hasText}", "INFO");
										if(hasText)
										{
											PPA.Core.Profiler.LogMessage($"处理形状 {i}，HasText=True", "INFO");
											var iShape = AdapterUtils.WrapShape(netApp, dynShape);
											if(iShape != null)
											{
												textFormatHelper.ApplyTextFormatting(iShape);
												hasFormatted=true;
												processedCount++;
											}
											else
											{
												PPA.Core.Profiler.LogMessage($"形状 {i} WrapShape 返回 null", "WARN");
											}
										}
									}
									else
									{
										PPA.Core.Profiler.LogMessage($"形状 {i} 获取失败", "WARN");
									}
								}
								catch(System.Exception ex3)
								{
									PPA.Core.Profiler.LogMessage($"处理形状 {i} 时出错: {ex3.Message}", "WARN");
								}
							}
							PPA.Core.Profiler.LogMessage($"dynamic 访问完成，共处理 {processedCount} 个形状", "INFO");
						}
						catch(System.Exception ex2)
						{
							PPA.Core.Profiler.LogMessage($"dynamic 访问也失败: {ex2.Message}", "ERROR");
							PPA.Core.Profiler.LogMessage($"堆栈跟踪: {ex2.StackTrace}", "ERROR");
						}
					}
				}

				if(hasFormatted)
				{
					Toast.Show(ResourceManager.GetString("Toast_FormatText_Success"),Toast.ToastType.Success);
				} else
				{
					Toast.Show(ResourceManager.GetString("Toast_FormatText_NoText"),Toast.ToastType.Warning);
				}
			});
		}

		// 适配包装逻辑已提取到 Core/Adapters/AdapterUtils.cs

		#endregion
	}
}

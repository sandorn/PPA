using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Logging;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Shape
{
	public static class MSOICrop
	{
		private static readonly ILogger _logger = LoggerProvider.GetLogger();

		/// <summary>
		/// MsoMergeCmd 到 idMso 命令的映射字典
		/// </summary>
		private static readonly Dictionary<MsoMergeCmd,string> MergeCmdToIdMso = new Dictionary<MsoMergeCmd,string>
		{
			{ MsoMergeCmd.msoMergeIntersect, OfficeCommands.ShapesIntersect },
			{ MsoMergeCmd.msoMergeUnion, OfficeCommands.ShapesUnion },
			{ MsoMergeCmd.msoMergeCombine, OfficeCommands.ShapesCombine },
			{ MsoMergeCmd.msoMergeSubtract, OfficeCommands.ShapesSubtract }
		};

		public static void CropShapesToSlide(NETOP.Application netApp)
		{
			netApp=ApplicationHelper.EnsureValidNetApplication(netApp);
			if(netApp==null)
			{
				_logger.LogError("无法获取 Application");
				return;
			}

			ExHandler.Run(() =>
			{
				var window = ExHandler.SafeGet(() => netApp.ActiveWindow, defaultValue:(NETOP.DocumentWindow)null);
				if(window==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				var view = ExHandler.SafeGet(() => window.View, defaultValue:(NETOP.View)null);
				var slide = view?.Slide as NETOP.Slide;
				var pageSetup = netApp.ActivePresentation?.PageSetup;
				float slideWidth = pageSetup?.SlideWidth ?? 0;
				float slideHeight = pageSetup?.SlideHeight ?? 0;
				if(slide==null||slideWidth<=0||slideHeight<=0)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlideSize"),Toast.ToastType.Warning);
					return;
				}

				var (shapesToCrop, collectionsToDispose)=CollectShapesToCrop(netApp,window.Selection,slide,slideWidth,slideHeight);
				if(shapesToCrop.Count==0)
				{
					Toast.Show(ResourceManager.GetString("Toast_CropShapes_None"),Toast.ToastType.Warning);
					return;
				}

				try
				{
					_logger.LogInformation($"开始裁剪 {shapesToCrop.Count} 个形状");
					foreach(var nativeShape in shapesToCrop)
					{
						if(nativeShape==null)
						{
							_logger.LogWarning("形状对象无效，跳过当前项");
							continue;
						}

						_logger.LogInformation($"裁剪形状: Id={nativeShape.Id}, Name={nativeShape.Name}");
						BooleanCrop(slide,nativeShape,slideWidth,slideHeight);
					}
				} finally
				{
					// 释放所有收集到的形状对象
					foreach(var shape in shapesToCrop)
					{
						try { shape.Dispose(); } catch { /* ignore */ }
					}
					// 释放集合对象
					collectionsToDispose.DisposeAll();
				}
			});
		}

		private static (List<NETOP.Shape> shapesToCrop, List<IDisposable> collectionsToDispose) CollectShapesToCrop(NETOP.Application app,NETOP.Selection selection,NETOP.Slide slide,float slideWidth,float slideHeight)
		{
			var shapesToCrop = new List<NETOP.Shape>();
			var collectionsToDispose = new List<IDisposable>();

			void TryAddShape(NETOP.Shape candidate)
			{
				if(candidate==null)
					return;
				if(!ShouldCropShape(candidate,slideWidth,slideHeight))
				{
					candidate.Dispose(); // 如果不需要裁剪，立即释放
					return;
				}

				shapesToCrop.Add(candidate);
			}

			if(selection!=null&&selection.Type==NetOffice.PowerPointApi.Enums.PpSelectionType.ppSelectionShapes)
			{
				var range = selection.ShapeRange;
				collectionsToDispose.Add(range);  // 延迟释放
				for(int i = 1;i<=range.Count;i++)
				{
					var shape = range[i];
					TryAddShape(shape);
				}
			} else
			{
				var slideShapes = slide.Shapes;
				collectionsToDispose.Add(slideShapes);  // 延迟释放
				foreach(NETOP.Shape shape in slideShapes)
				{
					TryAddShape(shape);
				}
			}
			return (shapesToCrop, collectionsToDispose);
		}

		private static void BooleanCrop(NETOP.Slide slide,NETOP.Shape target,float slideWidth,float slideHeight,MsoMergeCmd mergeCmd = MsoMergeCmd.msoMergeIntersect)
		{
			// 创建遮罩矩形
			NETOP.Shape mask = null;
			try
			{
				mask=slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,0,0,slideWidth,slideHeight);
				mask.Fill.Visible=MsoTriState.msoFalse;
				mask.Line.Visible=MsoTriState.msoFalse;

				// --- 步骤 1: 在运算前，保存主形状的 Z-Order 和幻灯片上所有形状的快照 ---
				int originalZOrder = target.ZOrderPosition;
				var beforeShapes = new HashSet<string>();
				using(var slideShapes1 = slide.Shapes)
				{
					foreach(NETOP.Shape s in slideShapes1)
					{
						using(s)
						{
							beforeShapes.Add($"{s.Id}|{s.Name}");
						}
					}
				}

				ExHandler.Run(() =>
				{
					// --- 步骤 2: 选中形状并执行布尔运算 ---
					var shapeNames = new string[] { target.Name, mask.Name };
					using(var shapeRange = slide.Shapes.Range((object) shapeNames))
					{
						shapeRange.Select();

						var version = slide.Application.Version;
						if(System.Version.Parse(version)>=System.Version.Parse("15.0"))
						{
							if(!MergeCmdToIdMso.TryGetValue(mergeCmd,out string idMso))
							{
								throw new ArgumentOutOfRangeException(nameof(mergeCmd),mergeCmd,"Unsupported merge command.");
							}

							_logger.LogInformation($"执行布尔运算命令: {idMso}");
							slide.Application.CommandBars.ExecuteMso(idMso);
						} else
						{
							_logger.LogWarning($"PowerPoint 版本 {version} 过低，不支持基于 idMso 的布尔运算。");
							Toast.Show(ResourceManager.GetString("Toast_OperationFailed_UnsupportedVersion"),Toast.ToastType.Error);
							return;
						}

						// 等待操作完成
						System.Threading.Thread.Sleep(100);

						// --- 步骤 3: 通过对比新旧形状列表，找到结果形状 ---
						string finalShapeName = null;
						using(var slideShapes2 = slide.Shapes)
						{
							foreach(NETOP.Shape shape in slideShapes2)
							{
								using(shape)
								{
									string key = $"{shape.Id}|{shape.Name}";
									if(!beforeShapes.Contains(key))
									{
										finalShapeName=shape.Name;
										break;
									}
								}
							}
						}

						if(!string.IsNullOrEmpty(finalShapeName))
						{
							// 重新获取结果形状，因为之前的引用已在 using 块中释放
							using(var finalShape = slide.Shapes[finalShapeName])
							{
								_logger.LogInformation($"找到结果形状: Id={finalShape.Id}, Name={finalShape.Name}，正在调整 Z-Order。");
								// --- 步骤 4: 调整结果形状的 Z-Order ---
								ExHandler.SafeSet(() =>
								{
									finalShape.ZOrder(MsoZOrderCmd.msoSendToBack);
									for(int i = 1;i<originalZOrder;i++)
										finalShape.ZOrder(MsoZOrderCmd.msoBringForward);
								});
							}
						} else
						{
							_logger.LogWarning("未能找到布尔运算的结果形状，无法调整 Z-Order。");
						}
					}
				});
			} catch
			{
				// 如果创建遮罩或执行布尔运算失败，确保释放遮罩
				if(mask!=null)
				{
					try
					{
						// 尝试释放遮罩（布尔运算后可能已被删除，忽略异常）
						mask.Dispose();
					} catch
					{
						// 忽略释放失败（对象可能已被删除）
					}
				}
				throw;
			} finally
			{
				// 确保遮罩被释放（布尔运算后可能已经被删除，这里只是保险措施）
				if(mask!=null)
				{
					try
					{
						// 尝试释放遮罩（如果已被删除，会抛出异常，忽略即可）
						mask.Dispose();
					} catch
					{
						// 忽略释放失败（对象可能已被删除）
					}
				}
			}
		}

		private static bool ShouldCropShape(NETOP.Shape shape,float slideWidth,float slideHeight)
		{
			if(shape==null)
				return false;

			if(shape.Type==MsoShapeType.msoPlaceholder||
				shape.Type==MsoShapeType.msoOLEControlObject||
				shape.Type==MsoShapeType.msoComment)
			{
				return false;
			}

			try
			{
				float left = shape.Left, top = shape.Top;
				float right = left + shape.Width, bottom = top + shape.Height;

				if(left<-0.5f||top<-0.5f||right>slideWidth+0.5f||bottom>slideHeight+0.5f)
				{
					return !(right<=0||bottom<=0||left>=slideWidth||top>=slideHeight);
				}

				return false;
			} catch
			{
				return false;
			}
		}
	}
}

/*
完整枚举值列表
枚举值	整数值	操作名称	功能描述
msoMergeUnion	0	联合	合并所有形状为一个整体，移除重叠边界
msoMergeCombine 1	组合	合并形状但保留重叠区域的边界
msoMergeFragment    2	拆分	将重叠区域分割为独立形状
msoMergeIntersect   3	相交	保留所有形状的重叠区域
msoMergeSubtract    4	剪除	从第一个形状中减去后续形
*/

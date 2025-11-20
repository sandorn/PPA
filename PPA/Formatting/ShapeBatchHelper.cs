using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using PPA.Core.Logging;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 形状批量操作辅助类
	/// </summary>
	internal class ShapeBatchHelper:IShapeBatchHelper
	{
		private readonly IShapeHelper _shapeHelper;
		private readonly ILogger _logger;

		public ShapeBatchHelper(IShapeHelper shapeHelper,ILogger logger = null)
		{
			_shapeHelper=shapeHelper??throw new ArgumentNullException(nameof(shapeHelper));
			_logger=logger??LoggerProvider.GetLogger();
		}

		/// <summary>
		/// 根据选中对象创建矩形外框：
		/// 1. 选中形状时：为每个形状创建包围框并考虑边框宽度
		/// 2. 选中幻灯片时：在每个幻灯片创建页面大小的矩形
		/// 3. 无选中时：在当前幻灯片创建页面大小的矩形
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		public void CreateBoundingBox(NETOP.Application app)
		{
			var netApp = app;

			ExHandler.Run(() =>
			{
				if(!TryRefreshContext(ref netApp,out var abstractApp))
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				var sel = GetSelectionWithRetry(ref netApp, ref abstractApp);
				if(sel==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoValidSelection"),Toast.ToastType.Warning);
					return;
				}

				var currentSlide = GetSlideWithRetry(ref netApp, ref abstractApp);
				if(currentSlide==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				UndoHelper.BeginUndoEntry(netApp,UndoHelper.UndoNames.CreateBoundingBox);

				// 获取幻灯片尺寸
				var pageSetup = netApp.ActivePresentation?.PageSetup;
				float slideWidth = pageSetup?.SlideWidth ?? 0;
				float slideHeight = pageSetup?.SlideHeight ?? 0;

				if(slideWidth<=0||slideHeight<=0)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlideSize"),Toast.ToastType.Warning);
					return;
				}

				List<NETOP.Shape> createdShapes = new List<NETOP.Shape>();
				string successMessage = "";

				// 1. 处理选中形状
				if(sel!=null)
				{
					// 处理单个形状
					if(sel is NETOP.Shape shape)
					{
						// 需要将 NETOP.Shape 转换为 IShape
						var abstractShape = AdapterUtils.WrapShape(netApp, shape);
						var (top, left, bottom, right)=_shapeHelper.GetShapeBorderWeights(abstractShape);

						// 计算矩形参数
						float rectLeft = shape.Left - left;
						float rectTop = shape.Top - top;
						float rectWidth = shape.Width + (left + right);
						float rectHeight = shape.Height + (top + bottom);

						// 创建矩形
						var abstractSlide = currentSlide as IComWrapper<NETOP.Slide>;
						if(abstractSlide!=null)
						{
							var rect = _shapeHelper.AddOneShape(currentSlide, rectLeft, rectTop, rectWidth, rectHeight, shape.Rotation);
							if(rect is IComWrapper<NETOP.Shape> rectWrapper)
							{
								createdShapes.Add(rectWrapper.NativeObject);
							}
						}
					}
					// 处理形状范围
					else if(sel is NETOP.ShapeRange shapes)
					{
						if(shapes.Count>0)
						{
							for(int i = 1;i<=shapes.Count;i++)
							{
								var currentShape = shapes[i];
								// 需要将 NETOP.Shape 转换为 IShape
								var abstractShape = AdapterUtils.WrapShape(netApp, currentShape);
								var (top, left, bottom, right)=_shapeHelper.GetShapeBorderWeights(abstractShape);

								// 计算矩形参数
								float rectLeft = currentShape.Left - left;
								float rectTop = currentShape.Top - top;
								float rectWidth = currentShape.Width + (left + right);
								float rectHeight = currentShape.Height + (top + bottom);

								// 创建矩形
								var rect = _shapeHelper.AddOneShape(currentSlide, rectLeft, rectTop, rectWidth, rectHeight, currentShape.Rotation);
								if(rect is IComWrapper<NETOP.Shape> rectWrapper)
								{
									createdShapes.Add(rectWrapper.NativeObject);
								}
							}
						}
					}

					if(createdShapes.Count>0)
					{
						var abstractSlide = currentSlide as IComWrapper<NETOP.Slide>;
						if(abstractSlide!=null)
						{
							var nativeSlide = abstractSlide.NativeObject;
							var shapeNames = createdShapes.Select(s => s.Name).ToArray();
							nativeSlide.Shapes.Range(shapeNames).Select();
						}
						successMessage=string.Format(ResourceManager.GetString("Toast_CreateBoundingBox_Shapes"),createdShapes.Count);
					}
				}
				// 2. 处理选中幻灯片 和 无选中
				else
				{
					// 创建要处理的幻灯片列表
					List<NETOP.Slide> slidesToProcess = new List<NETOP.Slide>();
					// 检查是否选中了幻灯片
					var window = app.ActiveWindow;
					if(window!=null&&window.Selection?.Type==NETOP.Enums.PpSelectionType.ppSelectionSlides)
					{
						// 选中幻灯片的情况
						var slideRange = window.Selection.SlideRange;
						if(slideRange?.Count>0)
						{
							for(int i = 1;i<=slideRange.Count;i++)
							{
								slidesToProcess.Add(slideRange[i]);
							}
							successMessage=string.Format(ResourceManager.GetString("Toast_CreateBoundingBox_Slides"),slideRange.Count);
						}
					} else
					{
						// 无选中的情况 - 需要从 ISlide 转换为 NETOP.Slide
						var abstractSlide = currentSlide as IComWrapper<NETOP.Slide>;
						if(abstractSlide!=null)
						{
							slidesToProcess.Add(abstractSlide.NativeObject);
						}
						successMessage=ResourceManager.GetString("Toast_CreateBoundingBox_PageSize");
					}

					// 统一处理幻灯片列表
					if(slidesToProcess.Count>0)
					{
						for(int i = 0;i<slidesToProcess.Count;i++)
						{
							var slide = slidesToProcess[i];
							// 将 NETOP.Slide 转换为 ISlide
							var abstractSlide = AdapterUtils.WrapSlide(netApp, slide);
							var rect = _shapeHelper.AddOneShape(abstractSlide, 0, 0, slideWidth, slideHeight);

							if(rect is IComWrapper<NETOP.Shape> rectWrapper)
							{
								var nativeRect = rectWrapper.NativeObject;
								createdShapes.Add(nativeRect);
								// 如果是第一张幻灯片，则选中其上的矩形
								if(i==0) nativeRect.Select();
							}
						}
					}
				}

				// 显示结果消息
				if(createdShapes.Count>0)
				{
					Toast.Show(successMessage,Toast.ToastType.Success);
				} else
				{
					Toast.Show(ResourceManager.GetString("Toast_CreateBoundingBox_None"),Toast.ToastType.Info);
				}
			});
		}

		/// <summary>
		/// 隐藏/显示对象：选中对象时隐藏选中对象，无选中对象时显示所有对象。
		/// </summary>
		/// <param name="netApp"> PowerPoint 应用程序实例。 </param>
		public void ToggleShapeVisibility(NETOP.Application netApp)
		{
			var currentApp = netApp;

			ExHandler.Run(() =>
			{
				if(!TryRefreshContext(ref currentApp,out var abstractApp))
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				var slide = GetSlideWithRetry(ref currentApp, ref abstractApp);
				if(slide==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				var selectionObj = GetSelectionWithRetry(ref currentApp, ref abstractApp);
				dynamic sel = selectionObj;
				if(sel!=null)
				{
					var shapesToHide = (sel switch
					{
						NETOP.ShapeRange range => range.Cast<NETOP.Shape>(),
						NETOP.Shape shape => new[] { shape },
						_ => Enumerable.Empty<NETOP.Shape>()
					}).ToList();

					if(shapesToHide.Count==0)
						return;

					UndoHelper.BeginUndoEntry(currentApp,UndoHelper.UndoNames.HideShapes);

					try
					{
						foreach(var shape in shapesToHide)
						{
							shape.Visible=MsoTriState.msoFalse;
						}

						var message = shapesToHide.Count == 1
							? ResourceManager.GetString("Toast_HideShapes_Single")
							: string.Format(ResourceManager.GetString("Toast_HideShapes_Multiple"), shapesToHide.Count);
						Toast.Show(message,Toast.ToastType.Success);
					} finally
					{
						shapesToHide.DisposeAll();
					}
				} else if(slide is IComWrapper<NETOP.Slide> { NativeObject: var nativeSlide })
				{
					ShowAllHiddenShapes(currentApp,nativeSlide.Shapes);
				}
			});
		}

		/// <summary>
		/// 显示幻灯片上所有被隐藏的形状。
		/// </summary>
		/// <param name="netApp"> PowerPoint 应用程序实例 </param>
		/// <param name="shapes"> 幻灯片的形状集合。 </param>
		private void ShowAllHiddenShapes(NETOP.Application netApp,NETOP.Shapes shapes)
		{
			List<NETOP.Shape> shapesToShow = new List<NETOP.Shape>();

			// 1. 找出所有需要显示的对象
			for(int i = 1;i<=shapes.Count;i++)
			{
				var shape = shapes[i];
				if(shape.Visible==MsoTriState.msoFalse)
				{
					shapesToShow.Add(shape);
				}
			}

			// 2. 根据列表内容执行操作和反馈
			if(shapesToShow.Count>0)
			{
				UndoHelper.BeginUndoEntry(netApp,UndoHelper.UndoNames.ShowShapes);
				try
				{
					foreach(var shape in shapesToShow)
					{
						shape.Visible=MsoTriState.msoTrue;
					}
					Toast.Show(string.Format(ResourceManager.GetString("Toast_ShowShapes_Multiple"),shapesToShow.Count),Toast.ToastType.Success);
				} finally
				{
					shapesToShow.DisposeAll();
				}
			} else
			{
				Toast.Show(ResourceManager.GetString("Toast_ShowShapes_None"),Toast.ToastType.Info);
			}
		}

		private bool TryRefreshContext(ref NETOP.Application netApp,out IApplication abstractApp)
		{
			netApp=ApplicationHelper.EnsureValidNetApplication(netApp);
			if(netApp==null)
			{
				abstractApp=null;
				return false;
			}

			abstractApp=ApplicationHelper.GetAbstractApplication(netApp);
			if(abstractApp==null)
			{
				_logger.LogWarning("无法获取抽象 Application");
				return false;
			}

			return true;
		}

		private dynamic GetSelectionWithRetry(ref NETOP.Application netApp,ref IApplication abstractApp)
		{
			var selection = _shapeHelper.ValidateSelection(abstractApp, showWarningWhenInvalid: false);
			if(selection!=null)
			{
				return selection;
			}

			_logger.LogWarning("返回 null，尝试刷新 Application 后重试");
			if(!TryRefreshContext(ref netApp,out abstractApp))
			{
				return null;
			}

			return _shapeHelper.ValidateSelection(abstractApp,showWarningWhenInvalid: false);
		}

		private ISlide GetSlideWithRetry(ref NETOP.Application netApp,ref IApplication abstractApp)
		{
			var slide = _shapeHelper.TryGetCurrentSlide(abstractApp);
			if(slide!=null)
			{
				return slide;
			}

			_logger.LogWarning("返回 null，尝试刷新 Application 后重试");
			if(!TryRefreshContext(ref netApp,out abstractApp))
			{
				return null;
			}

			return _shapeHelper.TryGetCurrentSlide(abstractApp);
		}
	}
}

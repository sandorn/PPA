using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using PPA.Core.Logging;
using PPA.Utilities;
using System;
using System.Runtime.InteropServices;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Shape
{
	/// <summary>
	/// 形状工具辅助类 提供形状相关的工具方法
	/// </summary>
	public class ShapeUtils:IShapeHelper
	{
		private readonly ILogger _logger = LoggerProvider.GetLogger();

		#region IShapeHelper 实现

		/// <summary>
		/// 创建单个矩形
		/// </summary>
		public IShape AddOneShape(ISlide slide,float left,float top,float width,float height,float rotation = 0)
		{
			if(slide==null) throw new ArgumentNullException(nameof(slide));

			// 使用 AdapterUtils 统一转换
			var netSlide = AdapterUtils.UnwrapSlide(slide);
			var netApp = AdapterUtils.UnwrapApplicationFromSlide(slide);

			if(netSlide==null||netApp==null) return null;

			var netShape = AddOneShapeInternal(netSlide, left, top, width, height, rotation);
			if(netShape!=null)
			{
				return AdapterUtils.WrapShape(netApp,netShape);
			}
			return null;
		}

		/// <summary>
		/// 获取形状的边框宽度
		/// </summary>
		public (float top, float left, float right, float bottom) GetShapeBorderWeights(IShape shape)
		{
			if(shape==null) return (0, 0, 0, 0);

			// 使用 AdapterUtils 统一转换
			var netShape = AdapterUtils.UnwrapShape(shape);

			if(netShape==null) return (0, 0, 0, 0);

			return GetShapeBorderWeightsInternal(netShape);
		}

		/// <summary>
		/// 检查一个COM对象是否已失效（例如，其底层应用程序已关闭）。
		/// </summary>
		/// <param name="comObj"> 要检查的COM对象。 </param>
		/// <returns> 如果对象为null或已失效，则为 true；否则为 false。 </returns>
		public bool IsInvalidComObject(object comObj)
		{
			// 1. 空对象或非COM对象，直接视为无效
			if(comObj==null||!Marshal.IsComObject(comObj))
			{
				return true;
			}

			// 2. 优先处理NetOffice包装的对象
			if(comObj is NetOffice.COMObject netOfficeObj)
			{
				// 检查 NetOffice 对象的 UnderlyingObject 是否为 null 如果底层对象已释放，UnderlyingObject 将为 null
				if(netOfficeObj.UnderlyingObject==null)
				{
					return true;
				}
				// 尝试访问 UnderlyingObject 来验证其有效性
				try
				{
					// 通过 Marshal.GetIUnknownForObject 检查底层对象是否有效
					Marshal.GetIUnknownForObject(netOfficeObj.UnderlyingObject);
					return false; // 成功获取，说明对象有效
				} catch
				{
					return true; // 底层对象已失效
				}
			}

			// 3. 对于原生COM对象，执行一个轻量级的"心跳"检测
			try
			{
				// Marshal.GetIUnknownForObject 会增加对象的引用计数。 如果对象已失效，此调用会抛出异常。 这是检查原生COM对象状态的一种非常可靠且快速的方法。
				Marshal.GetIUnknownForObject(comObj);
				return false; // 成功获取，说明对象有效
			} catch
			{
				// 任何异常都表明COM对象已不再可用
				return true;
			}
		}

		/// <summary>
		/// 尝试获取当前幻灯片
		/// </summary>
		public ISlide TryGetCurrentSlide(IApplication app)
		{
			if(app==null) return null;

			// 使用 AdapterUtils 统一转换
			var netApp = AdapterUtils.UnwrapApplication(app);

			if(netApp==null) return null;

			var currentApp = netApp;
			var netSlide = TryGetCurrentSlideInternal(ref currentApp);
			if(netSlide!=null)
			{
				// 如果内部刷新了 Application，需要使用最新实例包装
				return AdapterUtils.WrapSlide(currentApp,netSlide);
			}
			return null;
		}

		/// <summary>
		/// 验证并返回当前选择的对象
		/// </summary>
		public object ValidateSelection(IApplication app,bool requireMultipleShapes = false,bool showWarningWhenInvalid = true)
		{
			if(app==null) return null;

			// 使用 AdapterUtils 统一转换
			var netApp = AdapterUtils.UnwrapApplication(app);

			if(netApp==null) return null;

			return ValidateSelectionInternal(netApp,requireMultipleShapes,showWarningWhenInvalid);
		}

		#endregion IShapeHelper 实现

		#region 内部实现方法（使用具体类型）

		/// <summary>
		/// 创建单个矩形的内部实现
		/// </summary>
		private NETOP.Shape AddOneShapeInternal(NETOP.Slide slide,float left,float top,float width,float height,float rotation = 0)
		{
			if(slide==null) throw new ArgumentNullException(nameof(slide));
			if(width<=0||height<=0)
			{
				_logger.LogWarning($"无效尺寸: width={width}, height={height}");
				return null;
			}
			// 添加日志记录实际参数
			_logger.LogInformation($"创建矩形: L={left}, T={top}, W={width}, H={height}");

			return ExHandler.SafeGet(() =>
			{
				var rect = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
				// 隐藏矩形边框，确保无任何线条显示
				rect.Line.DashStyle=MsoLineDashStyle.msoLineSolid; // 实线，防止虚线样式影响
				rect.Line.Style=MsoLineStyle.msoLineSingle; // 确保线条样式为单线
				rect.Line.Weight=0;
				rect.Line.Transparency=1.0f; // 线条完全透明
				rect.Line.Visible=MsoTriState.msoFalse; // 确保线条不可见
				rect.Fill.Visible=MsoTriState.msoFalse; // 无填充
				rect.Top=top; rect.Left=left;//调整到合适位置

				rect.Rotation=rotation; // 如果需要旋转，可以设置角度
				return rect;
			},defaultValue: null);
		}

		/// <summary>
		/// 获取形状的边框宽度的内部实现
		/// </summary>
		private (float top, float left, float right, float bottom) GetShapeBorderWeightsInternal(NETOP.Shape shape)
		{
			return ExHandler.SafeGet(() =>
			{
				float top = 0, left = 0, right = 0, bottom = 0;

				if(shape.HasTable==MsoTriState.msoTrue)
				{
					var table = shape.Table;
					int rows = table.Rows.Count;
					int cols = table.Columns.Count;

					// 获取表格四个角的边框宽度
					top=(float) Math.Max(0,table.Cell(1,1).Borders[NETOP.Enums.PpBorderType.ppBorderTop].Weight);
					left=(float) Math.Max(0,table.Cell(1,1).Borders[NETOP.Enums.PpBorderType.ppBorderLeft].Weight);
					right=(float) Math.Max(0,table.Cell(rows,cols).Borders[NETOP.Enums.PpBorderType.ppBorderRight].Weight);
					bottom=(float) Math.Max(0,table.Cell(rows,cols).Borders[NETOP.Enums.PpBorderType.ppBorderBottom].Weight);
				} else if(shape.Line.Visible==MsoTriState.msoTrue)
				{
					// 普通形状使用统一的边框宽度
					top=left=right=bottom=(float) shape.Line.Weight;
				}
				return (top, left, right, bottom);
			},defaultValue: (0, 0, 0, 0));
		}

		/// <summary>
		/// 安全获取当前幻灯片的内部实现：通过 Interop 读取 SlideIndex，再通过 NetOffice 获取，避免直接访问 View.Slide 导致的本地化类名包装失败
		/// </summary>
		private NETOP.Slide TryGetCurrentSlideInternal(ref NETOP.Application netApp)
		{
			if(netApp==null)
			{
				_logger.LogWarning("netApp 为 null");
				return null;
			}

			NETOP.Application effectiveApp = netApp;
			NETOP.DocumentWindow window = ExHandler.SafeGet(() => effectiveApp.ActiveWindow, defaultValue:(NETOP.DocumentWindow)null);
			if(window==null)
			{
				window=TryRefreshActiveWindow(effectiveApp,out var refreshedApp);
				if(window==null)
				{
					_logger.LogWarning("无可用 ActiveWindow");
					return null;
				}
				effectiveApp=refreshedApp??effectiveApp;
				netApp=effectiveApp;
			}

			// 尝试通过 View.Slide 直接获取
			var slideViaView = ExHandler.SafeGet(() => window.View?.Slide as NETOP.Slide, defaultValue:(NETOP.Slide)null);
			if(slideViaView!=null)
			{
				return slideViaView;
			}

			// 备用：Selection.SlideRange
			var slideViaSelection = ExHandler.SafeGet(() =>
			{
				object selObj = window.Selection;
				dynamic sel = selObj;
				object srObj = sel?.SlideRange;
				dynamic sr = srObj;
				if(sr!=null && sr.Count>=1)
				{
					return sr[1] as NETOP.Slide;
				}
				return null;
			}, defaultValue:(NETOP.Slide)null);

			if(slideViaSelection==null)
			{
				_logger.LogWarning("无法通过 NetOffice 获取当前 Slide");
			}
			return slideViaSelection;
		}

		/// <summary>
		/// 验证并返回当前选择对象的内部实现
		/// </summary>
		private dynamic ValidateSelectionInternal(NETOP.Application app,bool requireMultipleShapes = false,bool showWarning = true)
		{
			if(app==null)
			{
				_logger.LogWarning("app 为 null");
				return null;
			}

			// 调试：检查 ActiveWindow
			var activeWindow = ExHandler.SafeGet(() => app.ActiveWindow, defaultValue: (NETOP.DocumentWindow)null);
			if(activeWindow==null)
			{
				activeWindow=TryRefreshActiveWindow(app,out app);
				if(activeWindow==null)
				{
					_logger.LogWarning("app.ActiveWindow 为 null，刷新后仍不可用");
					if(showWarning)
					{
						Toast.Show(ResourceManager.GetString("Toast_NoValidSelection"),Toast.ToastType.Warning);
					}
					return null;
				}
			}

			// 使用 object 类型避免 dynamic 在 lambda 中的类型推断问题
			var selectionObj = ExHandler.SafeGet(() => (object)(activeWindow.Selection), defaultValue: (object)null);
			dynamic selection = selectionObj;
			if(selection==null)
			{
				_logger.LogWarning("activeWindow.Selection 为 null");
				if(showWarning)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoValidSelection"),Toast.ToastType.Warning);
				}
				return null;
			}
			// --- 处理不同选择类型 ---
			switch(selection.Type)
			{
				case NETOP.Enums.PpSelectionType.ppSelectionShapes:
					// 检查是否需要多个形状
					int shapeCount = ExHandler.SafeGet(() => selection.ShapeRange?.Count ?? 0, defaultValue: 0);
					_logger.LogDebug($"选中类型=ppSelectionShapes, 形状数量={shapeCount}");

					if(requireMultipleShapes&&shapeCount<2)
					{
						if(showWarning)
						{
							Toast.Show(ResourceManager.GetString("Toast_NeedTwoShapes"),Toast.ToastType.Warning);
						}
						return null;
					}
					return selection.ShapeRange;

				case NETOP.Enums.PpSelectionType.ppSelectionText:
					// 在 NetOffice 中，无论是选中文本框还是光标在表格内，Type 都是 ppSelectionText 我们可以直接尝试获取包含它的 Shape，这个操作对两种情况都有效
					int textShapeCount = ExHandler.SafeGet(() => selection.ShapeRange?.Count ?? 0, defaultValue: 0);
					_logger.LogDebug($"选中类型=ppSelectionText, 形状数量={textShapeCount}");

					if(selection.ShapeRange!=null&&textShapeCount>0)
					{
						return selection.ShapeRange[1];
					}
					break;
			}

			// 如果所有情况都不匹配，则返回 null
			_logger.LogDebug("未匹配任何选择类型，返回 null");
			if(showWarning)
			{
				Toast.Show(ResourceManager.GetString("Toast_NoValidSelection"),Toast.ToastType.Warning);
			}
			return null;
		}

		/// <summary>
		/// 尝试刷新 ActiveWindow，当原应用对象失效时重新获取 NetOffice.Application
		/// </summary>
		private NETOP.DocumentWindow TryRefreshActiveWindow(NETOP.Application currentApp,out NETOP.Application refreshedApp)
		{
			refreshedApp=currentApp;
			var newApp = ApplicationHelper.GetNetOfficeApplication();
			if(newApp!=null&&!ReferenceEquals(newApp,currentApp))
			{
				refreshedApp=newApp;
				return ExHandler.SafeGet(() => newApp.ActiveWindow,defaultValue: (NETOP.DocumentWindow) null);
			}

			return null;
		}

		#endregion 内部实现方法（使用具体类型）
	}
}

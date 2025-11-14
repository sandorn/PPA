using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using PPA.Core.Adapters.PowerPoint;
using PPA.Utilities;
using System;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Shape
{
	/// <summary>
	/// 形状工具辅助类
	/// 提供形状相关的工具方法
	/// </summary>
	public class ShapeUtils : IShapeHelper
	{
		/// <summary>
		/// 静态实例，用于向后兼容（当无法从 DI 容器获取服务时使用）
		/// </summary>
		private static readonly ShapeUtils _defaultInstance = new ShapeUtils();

		/// <summary>
		/// 获取默认实例（用于向后兼容）
		/// </summary>
		public static ShapeUtils Default => _defaultInstance;

		#region Public Methods

		/// <summary>
		/// 创建单个矩形的辅助函数
		/// </summary>
		public NETOP.Shape AddOneShape(NETOP.Slide slide,float left,float top,float width,float height,float rotation = 0)
		{
			if(slide==null) throw new ArgumentNullException(nameof(slide));
			if(width<=0||height<=0)
			{
				Profiler.LogMessage($"[添加形状]无效尺寸: width={width}, height={height}");
				return null;
			}
			// 添加日志记录实际参数
			Profiler.LogMessage($"[添加形状]创建矩形: L={left}, T={top}, W={width}, H={height}");

			return ExHandler.Run(() =>
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
			},"[添加形状] 创建矩形");
		}

		/// <summary>
		/// 获取形状的边框宽度
		/// </summary>
		public (float top, float left, float right, float bottom) GetShapeBorderWeights(NETOP.Shape shape)
		{
			float top = 0, left = 0, right = 0, bottom = 0;

			ExHandler.Run(() =>
			{
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
			},"获取形状的边框宽度");
			return (top, left, right, bottom);
		}

		public bool IsInvalidComObject(object comObj)
		{
			// 简单方法检查对象状态
			if(comObj==null) return true;
			try
			{
				// 尝试访问对象的某个属性来验证其有效性 对于未知类型，我们不直接标记为无效，而是尝试进行类型安全的检查 对于NetOffice对象，我们可以尝试访问其基本属性
				if(comObj is NetOffice.COMObject netOfficeObj)
				{
					// 检查NetOffice对象是否有效
					if(netOfficeObj.UnderlyingObject==null) return true;
					// 对于特定类型，执行特定的验证
					switch(comObj)
					{
						case NETOP.Chart chart:
						{ var test = chart.Name; return false; }
						case NETOP.Axis axis:
						{ var test = axis.Type; return false; }
						default:
							// 对于其他类型的NetOffice对象，尝试安全地访问其属性来验证有效性
							try
							{
								// 尝试获取对象的Name属性
								var type = comObj.GetType();
								var nameProperty = type.GetProperty("Name");
								if(nameProperty!=null)
								{
									nameProperty.GetValue(comObj);
									return false;
								}

								// 如果没有Name属性，尝试获取Application属性
								var appProperty = type.GetProperty("Application");
								if(appProperty!=null)
								{
									appProperty.GetValue(comObj);
									return false;
								}
							} catch
							{
								// 如果属性访问失败，不立即标记为无效，让调用代码尝试操作 这样可以避免将有效的但属性访问方式不同的对象误判为无效
							}
							// 对于其他情况，默认认为对象是有效的，让调用代码尝试操作 这样可以避免将有效的但类型未知的COM对象误判为无效
							return false;
					}
				}
				// 默认情况下，我们假设对象是有效的，让调用代码尝试操作 如果对象确实无效，调用代码会捕获异常
				return false;
			} catch { return true; }
		}

		/// <summary>
		/// 安全获取当前幻灯片：通过 Interop 读取 SlideIndex，再通过 NetOffice 获取，避免直接访问 View.Slide 导致的本地化类名包装失败
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		/// <returns> 当前幻灯片对象，如果获取失败则返回 null </returns>
		public NETOP.Slide TryGetCurrentSlide(NETOP.Application app)
		{
			if(app==null) return null;
			try
			{
				// 优先通过 Interop 取索引，避免 NetOffice 包装本地化类名
				var underlying = (app as NetOffice.ICOMObject)?.UnderlyingObject as Microsoft.Office.Interop.PowerPoint.Application;
				int slideIndex = 0;
				try { slideIndex=underlying?.ActiveWindow?.View?.Slide?.SlideIndex??0; } catch(Exception ex) { Profiler.LogMessage($"TryGetCurrentSlide interop读取异常: {ex.Message}"); }

				if(slideIndex>0)
				{
					try { return app?.ActivePresentation?.Slides[slideIndex]; } catch(Exception ex) { Profiler.LogMessage($"TryGetCurrentSlide netoffice索引获取异常: {ex.Message}"); }
				}

				// 备选1：Selection.SlideRange
				try
				{
					var sel = app?.ActiveWindow?.Selection;
					var sr = sel?.SlideRange;
					if(sr!=null&&sr.Count>=1)
					{
						try { return sr[1]; } finally { sr?.Dispose(); }
					}
				} catch(Exception ex) { Profiler.LogMessage($"TryGetCurrentSlide选择范围异常: {ex.Message}"); }
			} catch(Exception ex) { Profiler.LogMessage($"TryGetCurrentSlide异常: {ex.Message}"); }
			return null;
		}

		/// <summary>
		/// 验证并返回当前选择的对象。
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例。 </param>
		/// <param name="requireMultipleShapes"> 是否要求必须选择多个形状。 </param>
		/// <returns>
		/// 返回一个动态对象，可能是：
		/// - ShapeRange (当选择多个形状时)
		/// - Shape (当选择单个形状、文本框或光标在表格内时)
		/// - null (如果选择无效或不满足条件)
		/// </returns>
		public dynamic ValidateSelection(NETOP.Application app,bool requireMultipleShapes = false)
		{
			// --- 安全检查 ---
			if(app?.ActiveWindow?.Selection==null)
			{
				Toast.Show(ResourceManager.GetString("Toast_NoValidSelection"),Toast.ToastType.Warning);
				return null;
			}

			var selection = app.ActiveWindow.Selection;

			// --- 处理不同选择类型 ---
			switch(selection.Type)
			{
				case NETOP.Enums.PpSelectionType.ppSelectionShapes:
					// 检查是否需要多个形状
					if(requireMultipleShapes&&(selection.ShapeRange?.Count??0)<2)
					{
						Toast.Show(ResourceManager.GetString("Toast_NeedTwoShapes"),Toast.ToastType.Warning);
						return null;
					}
					return selection.ShapeRange;

				case NETOP.Enums.PpSelectionType.ppSelectionText:
					// 在 NetOffice 中，无论是选中文本框还是光标在表格内，Type 都是 ppSelectionText 我们可以直接尝试获取包含它的 Shape，这个操作对两种情况都有效
					if(selection.ShapeRange!=null&&selection.ShapeRange.Count>0)
					{
						return selection.ShapeRange[1];
					}
					break;
			}

			// 如果所有情况都不匹配，则返回 null
			return null;
		}

		/// <summary>
		/// 抽象接口版本：获取当前幻灯片
		/// </summary>
		/// <param name="app">抽象应用实例</param>
		public ISlide TryGetCurrentSlide(IApplication app)
		{
			if(app==null) return null;

			if(app is IComWrapper<NETOP.Application> typed)
			{
				var native = TryGetCurrentSlide(typed.NativeObject);
				if(native!=null)
				{
					return AdapterUtils.WrapSlide(typed.NativeObject,native);
				}
			}

			if(app is IComWrapper wrapper)
			{
				if(wrapper.NativeObject is NETOP.Application netApp)
				{
					var native = TryGetCurrentSlide(netApp);
					if(native!=null)
					{
						return AdapterUtils.WrapSlide(netApp,native);
					}
				}
			}

			return null;
		}

		/// <summary>
		/// 抽象接口版本：验证当前选择
		/// </summary>
		public object ValidateSelection(IApplication app,bool requireMultipleShapes = false)
		{
			if(app==null) return null;

			if(app is IComWrapper<NETOP.Application> typed)
			{
				return ValidateSelection(typed.NativeObject,requireMultipleShapes);
			}

			if(app is IComWrapper wrapper)
			{
				if(wrapper.NativeObject is NETOP.Application netApp)
				{
					return ValidateSelection(netApp,requireMultipleShapes);
				}
			}

			return null;
		}


		#endregion Public Methods
	}
}

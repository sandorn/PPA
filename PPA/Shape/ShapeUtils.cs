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
		#region IShapeHelper 实现

		/// <summary>
		/// 创建单个矩形
		/// </summary>
		public IShape AddOneShape(ISlide slide, float left, float top, float width, float height, float rotation = 0)
		{
			if(slide==null) throw new ArgumentNullException(nameof(slide));

			// 转换为具体类型
			NETOP.Slide netSlide = null;
			NETOP.Application netApp = null;

			if(slide is IComWrapper<NETOP.Slide> typedSlide)
			{
				netSlide = typedSlide.NativeObject;
				if(slide.Application is IComWrapper<NETOP.Application> typedApp)
				{
					netApp = typedApp.NativeObject;
				}
			}
			else if(slide is IComWrapper wrapper)
			{
				if(wrapper.NativeObject is NETOP.Slide slideObj)
				{
					netSlide = slideObj;
					if(slide.Application is IComWrapper appWrapper && appWrapper.NativeObject is NETOP.Application appObj)
					{
						netApp = appObj;
					}
				}
			}

			if(netSlide==null || netApp==null) return null;

			var netShape = AddOneShapeInternal(netSlide, left, top, width, height, rotation);
			if(netShape!=null)
			{
				return AdapterUtils.WrapShape(netApp, netShape);
			}
			return null;
		}

		/// <summary>
		/// 获取形状的边框宽度
		/// </summary>
		public (float top, float left, float right, float bottom) GetShapeBorderWeights(IShape shape)
		{
			if(shape==null) return (0, 0, 0, 0);

			// 转换为具体类型
			NETOP.Shape netShape = null;

			if(shape is IComWrapper<NETOP.Shape> typed)
			{
				netShape = typed.NativeObject;
			}
			else if(shape is IComWrapper wrapper)
			{
				if(wrapper.NativeObject is NETOP.Shape shapeObj)
				{
					netShape = shapeObj;
				}
			}

			if(netShape==null) return (0, 0, 0, 0);

			return GetShapeBorderWeightsInternal(netShape);
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
		/// 尝试获取当前幻灯片
		/// </summary>
		public ISlide TryGetCurrentSlide(IApplication app)
		{
			if(app==null) return null;

			// 转换为具体类型
			NETOP.Application netApp = null;

			if(app is IComWrapper<NETOP.Application> typed)
			{
				netApp = typed.NativeObject;
			}
			else if(app is IComWrapper wrapper)
			{
				if(wrapper.NativeObject is NETOP.Application appObj)
				{
					netApp = appObj;
				}
			}

			if(netApp==null) return null;

			var netSlide = TryGetCurrentSlideInternal(netApp);
			if(netSlide!=null)
			{
				return AdapterUtils.WrapSlide(netApp, netSlide);
			}
			return null;
		}

		/// <summary>
		/// 验证并返回当前选择的对象
		/// </summary>
		public object ValidateSelection(IApplication app, bool requireMultipleShapes = false)
		{
			if(app==null) return null;

			// 转换为具体类型
			NETOP.Application netApp = null;

			if(app is IComWrapper<NETOP.Application> typed)
			{
				netApp = typed.NativeObject;
			}
			else if(app is IComWrapper wrapper)
			{
				if(wrapper.NativeObject is NETOP.Application appObj)
				{
					netApp = appObj;
				}
			}

			if(netApp==null) return null;

			return ValidateSelectionInternal(netApp, requireMultipleShapes);
		}

		#endregion

		#region 内部实现方法（使用具体类型）

		/// <summary>
		/// 创建单个矩形的内部实现
		/// </summary>
		private NETOP.Shape AddOneShapeInternal(NETOP.Slide slide, float left, float top, float width, float height, float rotation = 0)
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
		/// 获取形状的边框宽度的内部实现
		/// </summary>
		private (float top, float left, float right, float bottom) GetShapeBorderWeightsInternal(NETOP.Shape shape)
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

		/// <summary>
		/// 安全获取当前幻灯片的内部实现：通过 Interop 读取 SlideIndex，再通过 NetOffice 获取，避免直接访问 View.Slide 导致的本地化类名包装失败
		/// </summary>
		private NETOP.Slide TryGetCurrentSlideInternal(NETOP.Application netApp)
		{
			if(netApp==null) return null;
			try
			{
				// 优先通过 Interop 取索引，避免 NetOffice 包装本地化类名
				var nativeApp = ApplicationHelper.GetNativeComApplication(netApp);
				int slideIndex = 0;
				try { slideIndex=nativeApp.ActiveWindow.View.Slide.SlideIndex; } catch(Exception ex) { Profiler.LogMessage($"TryGetCurrentSlide interop读取异常: {ex.Message}"); }

				if(slideIndex>0)
				{
					try { return netApp.ActivePresentation.Slides[slideIndex]; } catch(Exception ex) { Profiler.LogMessage($"TryGetCurrentSlide netoffice索引获取异常: {ex.Message}"); }
				}

				// 备选1：Selection.SlideRange
				try
				{
					var sel = netApp.ActiveWindow.Selection;
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
		/// 验证并返回当前选择对象的内部实现
		/// </summary>
		private dynamic ValidateSelectionInternal(NETOP.Application app, bool requireMultipleShapes = false)
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

		#endregion
	}
}

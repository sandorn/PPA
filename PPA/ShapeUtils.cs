using NetOffice.OfficeApi.Enums;
using Project.Utilities;
using System;
using System.Diagnostics;
using ToastAPI;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Helpers
{
	public static class ShapeUtils
	{
		#region Public Methods

		/// <summary>
		/// 创建单个矩形的辅助函数
		/// </summary>
		public static NETOP.Shape AddOneShape(NETOP.Slide slide,float left,float top,float width,float height,float rotation = 0)
		{
			if(slide == null) throw new ArgumentNullException(nameof(slide));
			if(width <= 0 || height <= 0)
			{
				Debug.WriteLine($"[添加形状]无效尺寸: width={width}, height={height}");
				return null;
			}
			// 添加日志记录实际参数
			Debug.WriteLine($"[添加形状]创建矩形: L={left}, T={top}, W={width}, H={height}");

			return ExHandler.Run(() =>
			{
				var rect = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,left,top,width,height);
				// 隐藏矩形边框，确保无任何线条显示
				rect.Line.DashStyle = MsoLineDashStyle.msoLineSolid; // 实线，防止虚线样式影响
				rect.Line.Style = MsoLineStyle.msoLineSingle; // 确保线条样式为单线
				rect.Line.Weight = 0;
				rect.Line.Transparency = 1.0f; // 线条完全透明
				rect.Line.Visible = MsoTriState.msoFalse; // 确保线条不可见
				rect.Fill.Visible = MsoTriState.msoFalse; // 无填充
				rect.Top = top; rect.Left = left;//调整到合适位置

				rect.Rotation = rotation; // 如果需要旋转，可以设置角度
				return rect;
			},"[添加形状] 创建矩形");
		}

		/// <summary>
		/// 获取形状的边框宽度
		/// </summary>
		public static (float top, float left, float right, float bottom) GetShapeBorderWeights(NETOP.Shape shape)
		{
			float top = 0, left = 0, right = 0, bottom = 0;

			ExHandler.Run(() =>
			{
				if(shape.HasTable == MsoTriState.msoTrue)
				{
					var table = shape.Table;
					int rows = table.Rows.Count;
					int cols = table.Columns.Count;

					// 获取表格四个角的边框宽度
					top = (float) Math.Max(0,table.Cell(1,1).Borders[NETOP.Enums.PpBorderType.ppBorderTop].Weight);
					left = (float) Math.Max(0,table.Cell(1,1).Borders[NETOP.Enums.PpBorderType.ppBorderLeft].Weight);
					right = (float) Math.Max(0,table.Cell(rows,cols).Borders[NETOP.Enums.PpBorderType.ppBorderRight].Weight);
					bottom = (float) Math.Max(0,table.Cell(rows,cols).Borders[NETOP.Enums.PpBorderType.ppBorderBottom].Weight);
				} else if(shape.Line.Visible == MsoTriState.msoTrue)
				{
					// 普通形状使用统一的边框宽度
					top = left = right = bottom = (float) shape.Line.Weight;
				}
			},"获取形状的边框宽度");
			return (top, left, right, bottom);
		}

		public static bool IsInvalidComObject(object comObj)
		{
			// 简单方法检查对象状态
			if(comObj == null) return true;
			return ExHandler.Run(() =>
			{
				switch(comObj)
				{
					case NETOP.Chart chart:
					{
						var test = chart.Name;
						return false;
					}
					case NETOP.Axis axis:
					{
						var test = axis.Type;
						return false;
					}
					default:
						return true;
				}
			},"检查对象状态",defaultValue: true);
		}

		/// <summary>
		/// 跳过判断: 判断是否选取、选中是否为 shape、是否选中 2 个以上（可选参数）
		/// 不满足条件返回 true（跳过），满足条件返回 false（不跳过）
		/// </summary>
		/// <param name="app">PowerPoint 应用程序实例</param>
		/// <param name="requireMultipleShapes">可选：是否需要至少选择两个形状</param>
		public static bool ValidateSelection(NETOP.Application app,bool requireMultipleShapes = false)
		{
			var selection = app.ActiveWindow?.Selection;

			if(selection == null)
			{
				Toast.Show("未选择任何对象",Toast.ToastType.Warning);
				return true; // 验证失败
			}

			if(selection.Type != NETOP.Enums.PpSelectionType.ppSelectionShapes)
			{
				Toast.Show("请选选择形状对象",Toast.ToastType.Warning);
				return true; // 验证失败
			}

			if(requireMultipleShapes && (selection.ShapeRange?.Count ?? 0) < 2)
			{
				Toast.Show("请至少选择两个对象",Toast.ToastType.Warning);
				return true; // 验证失败
			}

			return false; // 验证通过
		}

		#endregion Public Methods
	}
}
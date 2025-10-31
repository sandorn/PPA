using NetOffice.OfficeApi.Enums;
using Project.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using ToastAPI;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Helpers
{
	/// <summary>
	/// 提供PowerPoint形状对齐、拉伸、吸附等相关操作的辅助方法。
	/// </summary>
	public static class AlignHelper
	{
		#region Public Methods

		// 下吸附：将第二个形状的下边与第一个形状的上边对齐，只移动第二个形状且只垂直移动
		public static void AttachBottom(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				var baseShape = shapes[1];   // 第一个选中的形状为基准
				var moveShape = shapes[2];   // 第二个选中的形状为要移动的对象

				// 只移动第二个形状的Top属性，使其下边与第一个形状的上边对齐
				moveShape.Top = baseShape.Top - moveShape.Height;

				Toast.Show("已将第二个对象下边与第一个对象上边对齐",Toast.ToastType.Success);
			},"上边对齐");
		}

		// 左吸附：将第二个形状的左边与第一个形状的右边对齐，只移动第二个形状且只水平移动
		public static void AttachLeft(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				var baseShape = shapes[1];   // 第一个选中的形状为基准
				var moveShape = shapes[2];   // 第二个选中的形状为要移动的对象

				// 只移动第二个形状的Left属性，使其左边与第一个形状的右边对齐
				moveShape.Left = baseShape.Left + baseShape.Width;

				Toast.Show("已将第二个对象左边与第一个对象右边对齐",Toast.ToastType.Success);
			},"右边对齐");
		}

		// 右吸附：将第二个形状的右边与第一个形状的左边对齐，只移动第二个形状且只水平移动
		public static void AttachRight(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				var baseShape = shapes[1];   // 第一个选中的形状为基准
				var moveShape = shapes[2];   // 第二个选中的形状为要移动的对象

				// 只移动第二个形状的Left属性，使其右边与第一个形状的左边对齐
				moveShape.Left = baseShape.Left - moveShape.Width;

				Toast.Show("已将第二个对象右边与第一个对象左边对齐",Toast.ToastType.Success);
			},"左边对齐");
		}

		// 上吸附：将第二个形状的上边与第一个形状的下边对齐，只移动第二个形状且只垂直移动
		public static void AttachTop(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				var baseShape = shapes[1];   // 第一个选中的形状为基准
				var moveShape = shapes[2];   // 第二个选中的形状为要移动的对象

				// 只移动第二个形状的Top属性，使其上边与第一个形状的下边对齐
				moveShape.Top = baseShape.Top + baseShape.Height;

				Toast.Show("已将第二个对象上边与第一个对象下边对齐",Toast.ToastType.Success);
			},"下边对齐");
		}

		// 底对齐到下方最近的水平参考线
		public static void GuideAlignBottom(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app)) return;
			var sel = app.ActiveWindow.Selection;

			ExHandler.Run(() =>
			{
				var guides = app.ActivePresentation.Guides;
				var horizontalGuides = new List<float>();
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation == NETOP.Enums.PpGuideOrientation.ppHorizontalGuide)
						horizontalGuides.Add(guide.Position);
				}
				if(horizontalGuides.Count == 0)
				{
					Toast.Show("当前文档没有水平参考线",Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in sel.ShapeRange)
				{
					float bottom = shape.Top + shape.Height;
					// 只考虑在下方的参考线
					var bottomGuides = horizontalGuides.Where(g => g >= bottom).ToList();
					if(bottomGuides.Count == 0) continue; // 没有下方参考线则跳过
					float nearest = bottomGuides[0];
					float minDist = Math.Abs(bottom - nearest);
					foreach(var guideY in bottomGuides)
					{
						float dist = Math.Abs(bottom - guideY);
						if(dist < minDist)
						{
							minDist = dist;
							nearest = guideY;
						}
					}
					shape.Top = nearest - shape.Height;
				}
				Toast.Show("已底对齐到参考线",Toast.ToastType.Success);
			},"底对齐到参考线");
		}

		// 水平居中到最近的两条垂直参考线的中点
		public static void GuideAlignHCenter(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app)) return;
			var sel = app.ActiveWindow.Selection;

			ExHandler.Run(() =>
			{
				var guides = app.ActivePresentation.Guides;
				var verticalGuides = new List<float>();
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation == NETOP.Enums.PpGuideOrientation.ppVerticalGuide)
						verticalGuides.Add(guide.Position);
				}
				if(verticalGuides.Count < 2)
				{
					Toast.Show("至少需要两条垂直参考线",Toast.ToastType.Warning);
					return;
				}
				verticalGuides.Sort();

				foreach(NETOP.Shape shape in sel.ShapeRange)
				{
					float center = shape.Left + (shape.Width / 2f);
					// 找到左侧最近的参考线a和右侧最近的参考线b
					float? a = null, b = null;
					foreach(var g in verticalGuides)
					{
						if(g <= center) a = g;
						if(g > center)
						{
							b = g;
							break;
						}
					}
					if(a == null || b == null) continue; // 没有包围中点的两条参考线则跳过

					float targetCenter = ((float) a + (float) b) / 2f;
					shape.Left = targetCenter - (shape.Width / 2f);
				}
				Toast.Show("已水平居中到参考线",Toast.ToastType.Success);
			},"水平居中到参考线");
		}

		// 左对齐到左侧最近的垂直参考线
		public static void GuideAlignLeft(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app)) return;
			var sel = app.ActiveWindow.Selection;

			ExHandler.Run(() =>
			{
				var guides = app.ActivePresentation.Guides;
				var verticalGuides = new List<float>();
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation == NETOP.Enums.PpGuideOrientation.ppVerticalGuide)
						verticalGuides.Add(guide.Position);
				}
				if(verticalGuides.Count == 0)
				{
					Toast.Show("当前文档没有垂直参考线",Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in sel.ShapeRange)
				{
					float left = shape.Left;
					// 只考虑在左侧的参考线
					var leftGuides = verticalGuides.Where(g => g <= left).ToList();
					if(leftGuides.Count == 0) continue; // 没有左侧参考线则跳过
					float nearest = leftGuides[0];
					float minDist = Math.Abs(left - nearest);
					foreach(var guideX in leftGuides)
					{
						float dist = Math.Abs(left - guideX);
						if(dist < minDist)
						{
							minDist = dist;
							nearest = guideX;
						}
					}
					shape.Left = nearest;
				}
				Toast.Show("已左对齐到参考线",Toast.ToastType.Success);
			},"左对齐到参考线");
		}

		// 右对齐到右侧最近的垂直参考线
		public static void GuideAlignRight(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app)) return;
			var sel = app.ActiveWindow.Selection;

			ExHandler.Run(() =>
			{
				var guides = app.ActivePresentation.Guides;
				var verticalGuides = new List<float>();
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation == NETOP.Enums.PpGuideOrientation.ppVerticalGuide)
						verticalGuides.Add(guide.Position);
				}
				if(verticalGuides.Count == 0)
				{
					Toast.Show("当前文档没有垂直参考线",Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in sel.ShapeRange)
				{
					float right = shape.Left + shape.Width;
					// 只考虑在右侧的参考线
					var rightGuides = verticalGuides.Where(g => g >= right).ToList();
					if(rightGuides.Count == 0) continue; // 没有右侧参考线则跳过
					float nearest = rightGuides[0];
					float minDist = Math.Abs(right - nearest);
					foreach(var guideX in rightGuides)
					{
						float dist = Math.Abs(right - guideX);
						if(dist < minDist)
						{
							minDist = dist;
							nearest = guideX;
						}
					}
					shape.Left = nearest - shape.Width;
				}
				Toast.Show("已右对齐到参考线",Toast.ToastType.Success);
			},"右对齐到参考线");
		}

		// 顶对齐到上方最近的水平参考线
		public static void GuideAlignTop(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app)) return;
			var sel = app.ActiveWindow.Selection;

			ExHandler.Run(() =>
			{
				var guides = app.ActivePresentation.Guides;
				var horizontalGuides = new List<float>();
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation == NETOP.Enums.PpGuideOrientation.ppHorizontalGuide)
						horizontalGuides.Add(guide.Position);
				}
				if(horizontalGuides.Count == 0)
				{
					Toast.Show("当前文档没有水平参考线",Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in sel.ShapeRange)
				{
					float top = shape.Top;
					// 只考虑在上方的参考线
					var topGuides = horizontalGuides.Where(g => g <= top).ToList();
					if(topGuides.Count == 0) continue; // 没有上方参考线则跳过
					float nearest = topGuides[0];
					float minDist = Math.Abs(top - nearest);
					foreach(var guideY in topGuides)
					{
						float dist = Math.Abs(top - guideY);
						if(dist < minDist)
						{
							minDist = dist;
							nearest = guideY;
						}
					}
					shape.Top = nearest;
				}
				Toast.Show("已顶对齐到参考线",Toast.ToastType.Success);
			},"顶对齐到参考线");
		}

		// 垂直居中到最近的两条水平参考线的中点
		public static void GuideAlignVCenter(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app)) return;
			var sel = app.ActiveWindow.Selection;

			ExHandler.Run(() =>
			{
				var guides = app.ActivePresentation.Guides;
				var horizontalGuides = new List<float>();
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation == NETOP.Enums.PpGuideOrientation.ppHorizontalGuide)
						horizontalGuides.Add(guide.Position);
				}
				if(horizontalGuides.Count < 2)
				{
					Toast.Show("至少需要两条水平参考线",Toast.ToastType.Warning);
					return;
				}
				horizontalGuides.Sort();

				foreach(NETOP.Shape shape in sel.ShapeRange)
				{
					float center = shape.Top + (shape.Height / 2f);
					// 找到上方最近的参考线a和下方最近的参考线b
					float? a = null, b = null;
					foreach(var g in horizontalGuides)
					{
						if(g <= center) a = g;
						if(g > center)
						{
							b = g;
							break;
						}
					}
					if(a == null || b == null) continue; // 没有包围中点的两条参考线则跳过

					float targetCenter = ((float) a + (float) b) / 2f;
					shape.Top = targetCenter - (shape.Height / 2f);
				}
				Toast.Show("已垂直居中到参考线",Toast.ToastType.Success);
			},"垂直居中到参考线");
		}

		// 高拉伸：高度拉伸到最近两条水平参考线之间并居中
		public static void GuidesStretchHeight(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app)) return;
			var sel = app.ActiveWindow.Selection;

			ExHandler.Run(() =>
			{
				var guides = app.ActivePresentation.Guides;
				var horizontalGuides = new List<float>();

				// 收集所有水平参考线
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation == NETOP.Enums.PpGuideOrientation.ppHorizontalGuide)
						horizontalGuides.Add(guide.Position);
				}

				// 检查参考线数量
				if(horizontalGuides.Count < 2)
				{
					Toast.Show("至少需要两条水平参考线",Toast.ToastType.Warning);
					return;
				}

				// 排序参考线位置（从上到下）
				horizontalGuides.Sort();

				// 处理每个选中形状
				foreach(NETOP.Shape shape in sel.ShapeRange)
				{
					// 计算形状垂直中心
					float centerY = shape.Top + (shape.Height / 2f);
					float? topGuide = null, bottomGuide = null;

					// 查找最近的上下参考线
					foreach(var guideY in horizontalGuides)
					{
						if(guideY <= centerY)
							topGuide = guideY;  // 当前参考线在中心上方
						if(guideY > centerY)
						{
							bottomGuide = guideY;  // 找到中心下方的第一条参考线
							break;
						}
					}

					// 应用参考线位置
					if(topGuide != null && bottomGuide != null)
					{
						shape.Top = (float) topGuide;
						shape.Height = (float) bottomGuide - (float) topGuide;
					}
				}

				Toast.Show("已将高度拉伸到参考线",Toast.ToastType.Success);
			},"高度拉伸到参考线");
		}

		// 宽高都拉伸：宽度和高度都拉伸到最近两条参考线之间并居中
		public static void GuidesStretchSize(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app)) return;
			var sel = app.ActiveWindow.Selection;

			ExHandler.Run(() =>
			{
				var guides = app.ActivePresentation.Guides;
				var verticalGuides = new List<float>();
				var horizontalGuides = new List<float>();

				// 一次性收集所有参考线
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation == NETOP.Enums.PpGuideOrientation.ppVerticalGuide)
						verticalGuides.Add(guide.Position);
					else if(guide.Orientation == NETOP.Enums.PpGuideOrientation.ppHorizontalGuide)
						horizontalGuides.Add(guide.Position);
				}

				// 检查参考线数量
				if(verticalGuides.Count < 2 || horizontalGuides.Count < 2)
				{
					string message = "";
					if(verticalGuides.Count < 2) message += "至少需要两条垂直参考线";
					if(horizontalGuides.Count < 2)
						message += (message != "" ? "和" : "") + "至少需要两条水平参考线";

					Toast.Show(message,Toast.ToastType.Warning);
					return;
				}

				// 排序参考线
				verticalGuides.Sort();
				horizontalGuides.Sort();

				// 处理每个选中形状
				foreach(NETOP.Shape shape in sel.ShapeRange)
				{
					// 处理高度
					float centerY = shape.Top + (shape.Height / 2f);
					float? topGuide = null, bottomGuide = null;
					foreach(var guideY in horizontalGuides)
					{
						if(guideY <= centerY) topGuide = guideY;
						if(guideY > centerY)
						{
							bottomGuide = guideY;
							break;
						}
					}
					if(topGuide != null && bottomGuide != null)
					{
						shape.Top = (float) topGuide;
						shape.Height = (float) bottomGuide - (float) topGuide;
					}

					// 处理宽度
					float centerX = shape.Left + (shape.Width / 2f);
					float? leftGuide = null, rightGuide = null;
					foreach(var guideX in verticalGuides)
					{
						if(guideX <= centerX) leftGuide = guideX;
						if(guideX > centerX)
						{
							rightGuide = guideX;
							break;
						}
					}
					if(leftGuide != null && rightGuide != null)
					{
						shape.Left = (float) leftGuide;
						shape.Width = (float) rightGuide - (float) leftGuide;
					}
				}

				Toast.Show("已将宽度和高度拉伸到参考线",Toast.ToastType.Success);
			},"高效双向拉伸");
		}

		// 宽拉伸：宽度拉伸到最近两条垂直参考线之间并居中
		public static void GuidesStretchWidth(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app)) return;
			var sel = app.ActiveWindow.Selection;

			ExHandler.Run(() =>
			{
				var guides = app.ActivePresentation.Guides;
				var verticalGuides = new List<float>();
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation == NETOP.Enums.PpGuideOrientation.ppVerticalGuide)
						verticalGuides.Add(guide.Position);
				}
				if(verticalGuides.Count < 2)
				{
					Toast.Show("至少需要两条垂直参考线",Toast.ToastType.Warning);
					return;
				}
				verticalGuides.Sort();

				foreach(NETOP.Shape shape in sel.ShapeRange)
				{
					float center = shape.Left + (shape.Width / 2f);
					float? a = null, b = null;
					foreach(var g in verticalGuides)
					{
						if(g <= center) a = g;
						if(g > center)
						{
							b = g;
							break;
						}
					}
					if(a == null || b == null) continue;
					shape.Left = (float) a;
					shape.Width = (float) b - (float) a;
				}
				Toast.Show("已将宽度拉伸到参考线",Toast.ToastType.Success);
			},"宽度拉伸到参考线");
		}

		// 设置选中对象等高
		public static void SetEqualHeight(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				float sourceHeight = shapes[1].Height;
				for(int i = 1;i <= shapes.Count;i++)
					shapes[i].Height = sourceHeight;
				Toast.Show("已设置等高",Toast.ToastType.Success);
			},"设置等高");
		}

		// 设置选中对象等宽且等高
		public static void SetEqualSize(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				float sourceWidth = shapes[1].Width, sourceHeight = shapes[1].Height;
				for(int i = 1;i <= shapes.Count;i++)
				{
					shapes[i].Width = sourceWidth;
					shapes[i].Height = sourceHeight;
				}
				Toast.Show("已设置等大小",Toast.ToastType.Success);
			},"设置等大小");
		}

		// 设置选中对象等宽
		public static void SetEqualWidth(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				float sourceWidth = shapes[1].Width;
				for(int i = 1;i <= shapes.Count;i++)
					shapes[i].Width = sourceWidth;
				Toast.Show("已设置等宽",Toast.ToastType.Success);
			},"设置等宽");
		}

		// 下延伸：下边对齐最下侧，上边位置保持不变（高度变大，上边不动）
		public static void StretchBottom(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				float maxBottom = float.MinValue;
				for(int i = 1;i <= shapes.Count;i++)
				{
					float bottom = shapes[i].Top + shapes[i].Height;
					if(bottom > maxBottom) maxBottom = bottom;
				}

				for(int i = 1;i <= shapes.Count;i++)
				{
					shapes[i].Height = maxBottom - shapes[i].Top;
				}
				Toast.Show("已向下延伸对齐",Toast.ToastType.Success);
			},"向下延伸对齐");
		}

		// 左延伸：左边对齐最左侧，右边位置保持不变（宽度变大，右边不动）
		public static void StretchLeft(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				float minLeft = float.MaxValue;
				for(int i = 1;i <= shapes.Count;i++)
					if(shapes[i].Left < minLeft) minLeft = shapes[i].Left;

				for(int i = 1;i <= shapes.Count;i++)
				{
					float right = shapes[i].Left + shapes[i].Width;
					shapes[i].Width = right - minLeft;
					shapes[i].Left = minLeft;
				}
				Toast.Show("已向左延伸对齐",Toast.ToastType.Success);
			},"向左延伸对齐");
		}

		// 右延伸：右边对齐最右侧，左边位置保持不变（宽度变大，左边不动）
		public static void StretchRight(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				float maxRight = float.MinValue;
				for(int i = 1;i <= shapes.Count;i++)
				{
					float right = shapes[i].Left + shapes[i].Width;
					if(right > maxRight) maxRight = right;
				}

				for(int i = 1;i <= shapes.Count;i++)
				{
					shapes[i].Width = maxRight - shapes[i].Left;
					// shapes[i].Left 不变
				}
				Toast.Show("已向右延伸对齐",Toast.ToastType.Success);
			},"向右延伸对齐");
		}

		// 上延伸：上边对齐最上侧，下边位置保持不变（高度变大，下边不动）
		public static void StretchTop(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				float minTop = float.MaxValue;
				for(int i = 1;i <= shapes.Count;i++)
					if(shapes[i].Top < minTop) minTop = shapes[i].Top;

				for(int i = 1;i <= shapes.Count;i++)
				{
					float bottom = shapes[i].Top + shapes[i].Height;
					shapes[i].Height = bottom - minTop;
					shapes[i].Top = minTop;
				}
				Toast.Show("已向上延伸对齐",Toast.ToastType.Success);
			},"向上延伸对齐");
		}

		// 交换两个选中对象的位置和大小
		public static void SwapSize(NETOP.Application app)
		{
			if(ShapeUtils.ValidateSelection(app,true)) return;
			var shapes = app.ActiveWindow.Selection.ShapeRange;

			ExHandler.Run(() =>
			{
				var shape1 = shapes[1];
				var shape2 = shapes[2];
				// 交换大小
				(shape1.Width, shape2.Width) = (shape2.Width, shape1.Width);
				(shape1.Height, shape2.Height) = (shape2.Height, shape1.Height);

				// 交换位置
				(shape1.Left, shape2.Left) = (shape2.Left, shape1.Left);
				(shape1.Top, shape2.Top) = (shape2.Top, shape1.Top);

				// 交换填充颜色
				(shape1.Fill.ForeColor.RGB, shape2.Fill.ForeColor.RGB) = (shape2.Fill.ForeColor.RGB, shape1.Fill.ForeColor.RGB);

				// 交换线条样式
				(shape1.Line.DashStyle, shape2.Line.DashStyle) = (shape2.Line.DashStyle, shape1.Line.DashStyle);
				(shape1.Line.Style, shape2.Line.Style) = (shape2.Line.Style, shape1.Line.Style);

				// 交换线条颜色
				(shape1.Line.ForeColor.RGB, shape2.Line.ForeColor.RGB) = (shape2.Line.ForeColor.RGB, shape1.Line.ForeColor.RGB);

				// 交换线条宽度
				(shape1.Line.Weight, shape2.Line.Weight) = (shape2.Line.Weight, shape1.Line.Weight);

				// 交换透明度
				(shape1.Fill.Transparency, shape2.Fill.Transparency) = (shape2.Fill.Transparency, shape1.Fill.Transparency);

				// 交换字体字号和颜色
				if(shape1.TextFrame.HasText == MsoTriState.msoTrue && shape2.TextFrame.HasText == MsoTriState.msoTrue)
				{
					var textRange1 = shape1.TextFrame.TextRange;
					var textRange2 = shape2.TextFrame.TextRange;

					// 交换字体字号
					(textRange1.Font.Size, textRange2.Font.Size) = (textRange2.Font.Size, textRange1.Font.Size);

					// 交换字体颜色
					(textRange1.Font.Color.RGB, textRange2.Font.Color.RGB) = (textRange2.Font.Color.RGB, textRange1.Font.Color.RGB);
				}
				Toast.Show("已交换大小和位置",Toast.ToastType.Success);
			},"交换大小和位置");
		}

		#endregion Public Methods
	}
}
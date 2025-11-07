using PPA.Core;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using MSOP = Microsoft.Office.Interop.PowerPoint;
using NETOP = NetOffice.PowerPointApi;
using Office = Microsoft.Office.Core;

namespace PPA.Shape
{
	public static class MSOICrop
	{
		#region Public Methods

		public static void CropShapesToSlide(NETOP.Application app)
		{
			MSOP.Application nativeApp = Globals.ThisAddIn.NativeApp;
			var NetApp = app;

			ExHandler.Run(() =>
			{
				var sel = nativeApp.ActiveWindow.Selection;
				var slide = nativeApp.ActiveWindow.View.Slide;

				float slideWidth = nativeApp.ActivePresentation.PageSetup.SlideWidth;
				float slideHeight = nativeApp.ActivePresentation.PageSetup.SlideHeight;

				// 获取要处理的形状列表
				var shapesToCrop = new List<MSOP.Shape>();

				// 情况1：有选中的形状
				if(sel!=null&&sel.Type==MSOP.PpSelectionType.ppSelectionShapes)
				{
					Profiler.LogMessage($"处理选中的 {sel.ShapeRange.Count} 个形状");

					for(int i = 1;i<=sel.ShapeRange.Count;i++)
					{
						var shape = sel.ShapeRange[i];
						if(ShouldCropShape(shape,slideWidth,slideHeight))
						{
							shapesToCrop.Add(shape);
						}
					}
				}
				// 情况2：没有选中的形状，处理幻灯片上所有形状
				else
				{
					Profiler.LogMessage("没有选中形状，处理幻灯片上所有形状");

					for(int i = 1;i<=slide.Shapes.Count;i++)
					{
						var shape = slide.Shapes[i];
						if(ShouldCropShape(shape,slideWidth,slideHeight))
						{
							shapesToCrop.Add(shape);
						}
					}
				}

				if(shapesToCrop.Count==0)
				{
					Toast.Show(ResourceManager.GetString("Toast_CropShapes_None"),Toast.ToastType.Warning);
					return;
				}

				Profiler.LogMessage($"开始裁剪 {shapesToCrop.Count} 个形状");

				// 逐个处理形状
				foreach(var shape in shapesToCrop)
				{
					Profiler.LogMessage($"裁剪形状: Id={shape.Id}, Name={shape.Name}");

					// 创建辅助矩形（无填充、无线条）
					MSOP.Shape rect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, slideWidth, slideHeight);

					// 设置辅助矩形为透明
					rect.Fill.Visible=Office.MsoTriState.msoFalse;
					rect.Line.Visible=Office.MsoTriState.msoFalse;

					// 执行裁剪
					BooleanCrop(slide,shape,rect);
				}
			});
		}

		// NetOffice PowerPoint API 裁剪形状到幻灯片范围，无效
		public static void CropShapesToSlideByNETOP(NETOP.Application app)
		{
			var slide = app.ActiveWindow.View.Slide as NETOP.Slide;
			var sel = app.ActiveWindow?.Selection;
			var presentation = app.ActivePresentation;
			var pageSetup = presentation?.PageSetup;
			float slideWidth = pageSetup?.SlideWidth ?? 0;
			float slideHeight = pageSetup?.SlideHeight ?? 0;

			List<NETOP.Shape> shapesToProcess = [];

			// 判断是否有选中形状
			if(sel!=null&&sel.Type==NETOP.Enums.PpSelectionType.ppSelectionShapes)
			{
				// 只处理选中的形状
				for(int i = 1;i<=sel.ShapeRange.Count;i++)
					shapesToProcess.Add(sel.ShapeRange[i]);
			} else
			{
				// 没有选中形状时，处理当前幻灯片上的所有形状
				for(int i = 1;i<=slide.Shapes.Count;i++)
					shapesToProcess.Add(slide.Shapes[i]);
			}

			foreach(var shape in shapesToProcess)
			{
				float left = shape.Left, top = shape.Top;
				float right = left + shape.Width, bottom = top + shape.Height;

				// 检查是否超出边界
				bool isOutside = left < -0.5f || top < -0.5f || right > slideWidth + 0.5f || bottom > slideHeight + 0.5f;
				if(!isOutside) continue;

				// 检查是否与幻灯片有重叠区域
				bool hasOverlap = !(right <= 0 || bottom <= 0 || left >= slideWidth || top >= slideHeight);
				if(!hasOverlap) continue;

				Profiler.LogMessage($"Processing: Id={shape.Id}, Name={shape.Name}, Type={shape.Type}");

				// 创建裁剪矩形
				NETOP.Shape slideRect = ShapeUtils.AddOneShape(slide, 0, 0, slideWidth, slideHeight);

				// 执行相交操作（ExHandler 会处理异常）
				ExHandler.Run(() =>
				{
					var shapeRange = slide.Shapes.Range(new object[] { shape.Name, slideRect.Name });
					Profiler.LogMessage($"shapeRange.Count={shapeRange.Count}");
					for(int i = 1;i<=shapeRange.Count;i++)
					{
						Profiler.LogMessage($"shapeRange[{i}].Name={shapeRange[i].Name}, Type={shapeRange[i].Type}");
					}
					shapeRange.MergeShapes(NetOffice.OfficeApi.Enums.MsoMergeCmd.msoMergeIntersect);

					// 如果操作成功，slideRect 会被 MergeShapes 操作处理，不需要手动删除
				});

				// 清理资源：如果操作失败，矩形可能仍然存在，尝试删除 注意：即使操作成功，尝试删除已不存在的对象也不会造成问题
				if(slideRect!=null)
				{
					try { slideRect.Delete(); } catch { /* 忽略：对象可能已被 MergeShapes 删除 */ }
				}
			}
		}

		#endregion Public Methods

		#region Private Methods

		private static void BooleanCrop(
			MSOP.Slide slide,MSOP.Shape shape1,
			MSOP.Shape slideRect,
			Office.MsoMergeCmd mergeCmd = Office.MsoMergeCmd.msoMergeIntersect)
		{
			// --- 辅助函数：安全地执行一个操作，忽略COM异常 ---
			static void SafeSet(Action action)
			{
				try
				{
					action();
				} catch(COMException ex) { Profiler.LogMessage($"异常信息: {ex.Message}"); }
			}

			// 存储需要恢复的属性
			int? originalRGBFill = null;
			int? originalRGBLine = null;
			float? originalLineWeight = null;
			int? originalTextRGB = null;
			bool hasText = false;
			int originalZOrder = shape1.ZOrderPosition;

			// 安全地获取原始形状属性
			SafeSet(() =>
			{
				if(shape1.Fill.Visible!=Office.MsoTriState.msoFalse)
					originalRGBFill=shape1.Fill.ForeColor.RGB;
			});

			SafeSet(() =>
			{
				if(shape1.Line.Visible!=Office.MsoTriState.msoFalse)
				{
					originalRGBLine=shape1.Line.ForeColor.RGB;
					originalLineWeight=shape1.Line.Weight;
				}
			});

			SafeSet(() =>
			{
				if(shape1.TextFrame2.HasText==Office.MsoTriState.msoTrue)
				{
					originalTextRGB=shape1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB;
					hasText=true;
				}
			});

			// 确保原始形状是最后一个被选中的
			slideRect.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
			shape1.ZOrder(Office.MsoZOrderCmd.msoBringToFront);

			// 记录合并前形状标识
			var beforeShapes = new HashSet<string>();
			foreach(MSOP.Shape shape in slide.Shapes)
			{
				beforeShapes.Add($"{shape.Id}|{shape.Name}");
			}

			MSOP.Shape newShape = null;

			ExHandler.Run(() =>
			{
				// 创建形状范围并执行合并操作
				var shapeRange = slide.Shapes.Range(new object[] { slideRect.Name, shape1.Name });
				shapeRange.MergeShapes(mergeCmd);

				// 查找新生成的形状
				foreach(MSOP.Shape shape in slide.Shapes)
				{
					string key = $"{shape.Id}|{shape.Name}";
					if(!beforeShapes.Contains(key))
					{
						newShape=shape;
						break;
					}
				}

				// 安全地恢复原始属性到新形状
				if(newShape!=null)
				{
					// 恢复填充色
					if(originalRGBFill.HasValue)
						SafeSet(() => newShape.Fill.ForeColor.RGB=originalRGBFill.Value);

					// 恢复轮廓
					if(originalRGBLine.HasValue)
						SafeSet(() => newShape.Line.ForeColor.RGB=originalRGBLine.Value);

					if(originalLineWeight.HasValue)
					{
						float weightToSet = originalLineWeight.Value;
						if(weightToSet<=0) weightToSet=1.0f; // 修复：确保线条粗细有效
						SafeSet(() => newShape.Line.Weight=weightToSet);
					}

					// 恢复文本格式
					if(hasText&&originalTextRGB.HasValue)
						SafeSet(() => newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB=originalTextRGB.Value);

					// 恢复Z轴顺序
					SafeSet(() =>
					{
						newShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
						for(int i = 1;i<originalZOrder;i++)
							newShape.ZOrder(Office.MsoZOrderCmd.msoBringForward);
					});
				}
			});
		}

		#endregion Private Methods

		#region Private Methods

		// 辅助方法：判断形状是否需要裁剪
		private static bool ShouldCropShape(MSOP.Shape shape,float slideWidth,float slideHeight)
		{
			// 排除占位符、OLE控件、批注等不需要裁剪的形状类型
			if(shape.Type==Office.MsoShapeType.msoPlaceholder||
				shape.Type==Office.MsoShapeType.msoOLEControlObject||
				shape.Type==Office.MsoShapeType.msoComment)
			{
				return false;
			}

			try
			{
				float left = shape.Left, top = shape.Top;
				float right = left + shape.Width, bottom = top + shape.Height;

				// 检查是否超出幻灯片边界
				if(left<-0.5f||top<-0.5f||right>slideWidth+0.5f||bottom>slideHeight+0.5f)
				{
					// 检查是否与幻灯片有重叠区域
					return !(right<=0||bottom<=0||left>=slideWidth||top>=slideHeight);
				} else
				{
					// 如果不超出幻灯片边界，直接返回 false
					return false;
				}
			} catch(COMException)
			{
				// 处理可能出现的COM异常（如形状已被删除）
				return false;
			}
		}

		#endregion Private Methods
	}

	public static class ShapeForUtils
	{
		#region Public Methods

		/// <summary>
		/// 将源形状的格式（样式）应用到目标形状，实现类似“格式刷”的功能。 仅当两者类型一致时才执行，避免样式应用异常。
		/// </summary>
		/// <param name="sourceShape"> 格式来源形状。 </param>
		/// <param name="targetShape"> 要应用格式的目标形状。 </param>
		public static void ApplyShapeFormat(MSOP.Shape sourceShape,MSOP.Shape targetShape)
		{
			// 判断类型是否一致
			if(sourceShape.Type==targetShape.Type&&
				sourceShape.AutoShapeType==targetShape.AutoShapeType)
			{
				sourceShape.PickUp();
				targetShape.Apply();
			} else
			{
				// 可根据需要抛出异常或记录日志
				Profiler.LogMessage("源形状与目标形状类型不一致，未执行格式刷。");
			}
		}

		/// <summary>
		/// 使用 AddShape 方法手动创建形状副本。 适用于简单形状的复制，需手动复制位置、大小、填充、线条等属性。
		/// </summary>
		/// <param name="slide"> 目标幻灯片对象。 </param>
		/// <param name="shape"> 要复制的形状对象。 </param>
		/// <returns> 复制后的新形状对象。 </returns>
		/// 调用方法：var newShape = ShapeUtils.CopyShapeUsingAddShape(slide, shape1);
		public static MSOP.Shape CopyShapeUsingAddShape(MSOP.Slide slide,MSOP.Shape shape)
		{
			var copyShape = slide.Shapes.AddShape(
				shape.AutoShapeType,
				shape.Left,
				shape.Top,
				shape.Width,
				shape.Height
			);

			// 复制填充、线条等属性
			copyShape.Fill.ForeColor.RGB=shape.Fill.ForeColor.RGB;
			copyShape.Line.ForeColor.RGB=shape.Line.ForeColor.RGB;
			copyShape.Rotation=shape.Rotation;

			return copyShape;
		}

		/// <summary>
		/// 使用 Group 和 Ungroup 方法复制复杂形状。 适用于复杂形状的复制，但代码稍显繁琐。
		/// </summary>
		/// <param name="slide"> 目标幻灯片对象。 </param>
		/// <param name="shape"> 要复制的形状对象。 </param>
		/// <returns> 复制后的新形状对象。 </returns>
		public static MSOP.Shape CopyShapeUsingGroupAndUngroup(MSOP.Slide slide,MSOP.Shape shape)
		{
			var tempShape = slide.Shapes.AddShape(
				Office.MsoAutoShapeType.msoShapeRectangle,
				0, 0, 1, 1
			);

			var group = slide.Shapes.Range(new object[] { shape.Name, tempShape.Name }).Group();
			var copyShape = group.Ungroup()[1];

			// 删除临时形状
			tempShape.Delete();

			return copyShape;
		}

		/// <summary>
		/// 使用 PickUp 和 Apply 方法复制形状样式。 适用于快速复制样式，但不会复制内容（如文本或图片）。
		/// </summary>
		/// <param name="slide"> 目标幻灯片对象。 </param>
		/// <param name="shape"> 要复制的形状对象。 </param>
		/// <returns> 复制后的新形状对象。 </returns>
		public static MSOP.Shape CopyShapeUsingPickUpAndApply(MSOP.Slide slide,MSOP.Shape shape)
		{
			var copyShape = slide.Shapes.AddShape(
				shape.AutoShapeType,
				shape.Left,
				shape.Top,
				shape.Width,
				shape.Height
			);

			// 复制样式
			shape.PickUp();
			copyShape.Apply();

			return copyShape;
		}

		#endregion Public Methods
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

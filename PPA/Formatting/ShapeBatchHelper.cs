using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Shape;
using PPA.Utilities;
using System.Collections.Generic;
using System.Linq;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 形状批量操作辅助类
	/// </summary>
	public static class ShapeBatchHelper
	{
		/// <summary>
		/// 根据选中对象创建矩形外框：
		/// 1. 选中形状时：为每个形状创建包围框并考虑边框宽度
		/// 2. 选中幻灯片时：在每个幻灯片创建页面大小的矩形
		/// 3. 无选中时：在当前幻灯片创建页面大小的矩形
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		public static void Bt601_Click(NETOP.Application app)
		{
			UndoHelper.BeginUndoEntry(app,UndoHelper.UndoNames.CreateBoundingBox);

			ExHandler.Run(() =>
			{
				var sel = ShapeUtils.ValidateSelection(app);
				var currentSlide = ShapeUtils.TryGetCurrentSlide(app);

				if(currentSlide==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				// 获取幻灯片尺寸
				var pageSetup = app.ActivePresentation?.PageSetup;
				float slideWidth = pageSetup?.SlideWidth ?? 0;
				float slideHeight = pageSetup?.SlideHeight ?? 0;

				if(slideWidth<=0||slideHeight<=0)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlideSize"),Toast.ToastType.Warning);
					return;
				}

				List<NETOP.Shape> createdShapes = [];
				string successMessage = "";

				// 1. 处理选中形状
				if(sel!=null)
				{
					// 处理单个形状
					if(sel is NETOP.Shape shape)
					{
						var (top, left, bottom, right)=ShapeUtils.GetShapeBorderWeights(shape);

						// 计算矩形参数
						float rectLeft = shape.Left - left;
						float rectTop = shape.Top - top;
						float rectWidth = shape.Width + (left + right);
						float rectHeight = shape.Height + (top + bottom);

						// 创建矩形
						var rect = ShapeUtils.AddOneShape(currentSlide, rectLeft, rectTop, rectWidth, rectHeight, shape.Rotation);
						if(rect!=null) createdShapes.Add(rect);
					}
					// 处理形状范围
					else if(sel is NETOP.ShapeRange shapes)
					{
						if(shapes.Count>0)
						{
							for(int i = 1;i<=shapes.Count;i++)
							{
								var currentShape = shapes[i];
								var (top, left, bottom, right)=ShapeUtils.GetShapeBorderWeights(currentShape);

								// 计算矩形参数
								float rectLeft = currentShape.Left - left;
								float rectTop = currentShape.Top - top;
								float rectWidth = currentShape.Width + (left + right);
								float rectHeight = currentShape.Height + (top + bottom);

								// 创建矩形
								var rect = ShapeUtils.AddOneShape(currentSlide, rectLeft, rectTop, rectWidth, rectHeight, currentShape.Rotation);

								if(rect!=null) createdShapes.Add(rect);
							}
						}
					}

					if(createdShapes.Count>0)
					{
						var shapeNames = createdShapes.Select(s => s.Name).ToArray();
						currentSlide.Shapes.Range(shapeNames).Select();
						successMessage=ResourceManager.GetString("Toast_CreateBoundingBox_Shapes",createdShapes.Count);
					}
				}
				// 2. 处理选中幻灯片 和 无选中
				else
				{
					// 创建要处理的幻灯片列表
					List<NETOP.Slide> slidesToProcess = [];
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
							successMessage=ResourceManager.GetString("Toast_CreateBoundingBox_Slides",slideRange.Count);
						}
					} else
					{
						// 无选中的情况
						slidesToProcess.Add(currentSlide);
						successMessage=ResourceManager.GetString("Toast_CreateBoundingBox_PageSize");
					}

					// 统一处理幻灯片列表
					if(slidesToProcess.Count>0)
					{
						for(int i = 0;i<slidesToProcess.Count;i++)
						{
							var slide = slidesToProcess[i];
							var rect = ShapeUtils.AddOneShape(slide, 0, 0, slideWidth, slideHeight);

							if(rect!=null)
							{
								createdShapes.Add(rect);
								// 如果是第一张幻灯片，则选中其上的矩形
								if(i==0) rect.Select();
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
		/// <param name="app"> PowerPoint 应用程序实例。 </param>
		public static void ToggleShapeVisibility(NETOP.Application app)
		{
			ExHandler.Run(() =>
			{
				var slide = ShapeUtils.TryGetCurrentSlide(app);
				if(slide==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				var sel = ShapeUtils.ValidateSelection(app);
				if(sel!=null)
				{
					// --- 场景1: 隐藏选中的对象 ---
					if(sel is NETOP.Shape shape)
					{
						// 单个形状的情况，创建临时ShapeRange
						List<NETOP.Shape> shapeList = [shape];
						UndoHelper.BeginUndoEntry(app,UndoHelper.UndoNames.HideShapes);
						try
						{
							shape.Visible=MsoTriState.msoFalse;
							Toast.Show(ResourceManager.GetString("Toast_HideShapes_Single"),Toast.ToastType.Success);
						} finally
						{
							shapeList.DisposeAll();
						}
					} else if(sel is NETOP.ShapeRange shapeRange)
					{
						// 多个形状的情况
						HideSelectedShapes(app,shapeRange);
					}
				} else
				{
					// --- 场景2: 显示所有对象 ---
					ShowAllHiddenShapes(app,slide.Shapes);
				}
			});
		}

		/// <summary>
		/// 隐藏指定形状范围内的所有形状。
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		/// <param name="shapeRange"> 要隐藏的形状范围。 </param>
		private static void HideSelectedShapes(NETOP.Application app,NETOP.ShapeRange shapeRange)
		{
			// 使用目标类型 new() 和集合表达式 [] (C# 9.0+ & C# 12.0)
			List<NETOP.Shape> shapesToHide = new(shapeRange.Count);
			for(int i = 1;i<=shapeRange.Count;i++)
			{
				shapesToHide.Add(shapeRange[i]);
			}

			UndoHelper.BeginUndoEntry(app,UndoHelper.UndoNames.HideShapes);
			try
			{
				foreach(var shape in shapesToHide)
				{
					shape.Visible=MsoTriState.msoFalse;
				}
				Toast.Show(ResourceManager.GetString("Toast_HideShapes_Multiple",shapesToHide.Count),Toast.ToastType.Success);
			} finally
			{
				shapesToHide.DisposeAll();
			}
		}

		/// <summary>
		/// 显示幻灯片上所有被隐藏的形状。
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		/// <param name="shapes"> 幻灯片的形状集合。 </param>
		private static void ShowAllHiddenShapes(NETOP.Application app,NETOP.Shapes shapes)
		{
			List<NETOP.Shape> shapesToShow = [];

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
				UndoHelper.BeginUndoEntry(app,UndoHelper.UndoNames.ShowShapes);
				try
				{
					foreach(var shape in shapesToShow)
					{
						shape.Visible=MsoTriState.msoTrue;
					}
					Toast.Show(ResourceManager.GetString("Toast_ShowShapes_Multiple",shapesToShow.Count),Toast.ToastType.Success);
				} finally
				{
					shapesToShow.DisposeAll();
				}
			} else
			{
				Toast.Show(ResourceManager.GetString("Toast_ShowShapes_None"),Toast.ToastType.Info);
			}
		}
	}
}

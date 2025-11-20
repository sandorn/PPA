using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Logging;
using PPA.Shape;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using static PPA.Core.Abstraction.Business.OfficeCommands;
using AlignmentType = PPA.Core.Abstraction.Business.AlignmentType;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 提供PowerPoint形状对齐、拉伸、吸附等相关操作的辅助方法。
	/// </summary>
	public class AlignHelper:IAlignHelper
	{
		private readonly IShapeHelper _shapeHelper;
		private readonly ICommandExecutor _commandExecutor;
		private readonly ILogger _logger;

		// 缓存适配器对象，避免重复转换，提升性能
		private IApplication _cachedAbstractApp; // 缓存的抽象接口

		private NETOP.Application _cachedNetApp; // 缓存的 NetOffice 对象

		/// <summary>
		/// 构造函数，通过依赖注入获取服务
		/// </summary>
		/// <param name="shapeHelper"> 形状工具服务（可选，如果为 null 则创建新实例） </param>
		/// <param name="commandExecutor"> 命令执行器（可选，如果为 null 则从 DI 容器获取） </param>
		/// <param name="logger"> 日志记录器（可选，如果为 null 则使用默认日志记录器） </param>
		public AlignHelper(IShapeHelper shapeHelper = null,ICommandExecutor commandExecutor = null,ILogger logger = null)
		{
			var provider = ApplicationProvider.Current;

			if(shapeHelper==null)
			{
				shapeHelper=provider?.ServiceProvider?.GetService(typeof(IShapeHelper)) as IShapeHelper;
			}
			_shapeHelper=shapeHelper??new ShapeUtils();

			if(commandExecutor==null)
			{
				commandExecutor=provider?.ServiceProvider?.GetService(typeof(ICommandExecutor)) as ICommandExecutor;
			}
			_commandExecutor=commandExecutor;

			_logger=logger??LoggerProvider.GetLogger();
		}

		public void GuidesStretchHeight(IApplication abstractApp) => InvokeWithNative(abstractApp,GuidesStretchHeight);

		public void GuideAlignBottom(IApplication abstractApp) => InvokeWithNative(abstractApp,GuideAlignBottom);

		public void StretchBottom(IApplication abstractApp) => InvokeWithNative(abstractApp,StretchBottom);

		public void SwapSize(IApplication abstractApp) => InvokeWithNative(abstractApp,SwapSize);

		public void SetEqualHeight(IApplication abstractApp) => InvokeWithNative(abstractApp,SetEqualHeight);

		public void GuidesStretchSize(IApplication abstractApp) => InvokeWithNative(abstractApp,GuidesStretchSize);

		public void GuideAlignHCenter(IApplication abstractApp) => InvokeWithNative(abstractApp,GuideAlignHCenter);

		public void StretchLeft(IApplication abstractApp) => InvokeWithNative(abstractApp,StretchLeft);

		/// <summary>
		/// 将抽象接口转换为 NetOffice 对象，带缓存优化
		/// </summary>
		/// <param name="abstractApp"> 抽象应用程序接口 </param>
		/// <param name="action"> 需要执行的操作 </param>
		private void InvokeWithNative(IApplication abstractApp,Action<NETOP.Application> action)
		{
			if(action==null) return;

			// 使用缓存避免重复转换
			NETOP.Application netApp;
			if(_cachedAbstractApp==abstractApp&&_cachedNetApp!=null)
			{
				netApp=_cachedNetApp;
			} else
			{
				netApp=ApplicationHelper.GetNetOfficeApplication(abstractApp);
				if(netApp!=null)
				{
					_cachedAbstractApp=abstractApp;
					_cachedNetApp=netApp;
				}
			}

			if(netApp==null) return;
			action(netApp);
		}

		/// <summary>
		/// 将 NetOffice Application 转换为抽象接口，带缓存优化
		/// </summary>
		/// <param name="netApp"> NetOffice 应用程序对象 </param>
		/// <returns> 抽象接口对象 </returns>
		private IApplication GetAbstractApp(NETOP.Application netApp)
		{
			if(netApp==null) return null;

			// 如果缓存的 NetOffice 对象匹配，直接返回缓存的抽象接口
			if(_cachedNetApp==netApp&&_cachedAbstractApp!=null)
			{
				return _cachedAbstractApp;
			}

			// 使用 ApplicationHelper 进行转换（统一管理转换逻辑）
			var abstractApp = ApplicationHelper.GetAbstractApplication(netApp);
			_cachedNetApp=netApp;
			_cachedAbstractApp=abstractApp;
			return abstractApp;
		}

		public void GuidesStretchWidth(IApplication abstractApp) => InvokeWithNative(abstractApp,GuidesStretchWidth);

		public void GuideAlignLeft(IApplication abstractApp) => InvokeWithNative(abstractApp,GuideAlignLeft);

		public void StretchRight(IApplication abstractApp) => InvokeWithNative(abstractApp,StretchRight);

		#region Public Methods

		// 下吸附：将第二个形状的下边与第一个形状的上边对齐，只移动第二个形状且只垂直移动
		public void AttachBottom(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				var baseShape = shapes[1];   // 第一个选中的形状为基准
				var moveShape = shapes[2];   // 第二个选中的形状为要移动的对象

				// 只移动第二个形状的Top属性，使其下边与第一个形状的上边对齐
				moveShape.Top=baseShape.Top-moveShape.Height;

				Toast.Show(ResourceManager.GetString("Align_AttachBottom_Success","已将第二个对象下边与第一个对象上边对齐"),Toast.ToastType.Success);
			});
		}

		public void GuideAlignRight(IApplication abstractApp) => InvokeWithNative(abstractApp,GuideAlignRight);

		public void StretchTop(IApplication abstractApp) => InvokeWithNative(abstractApp,StretchTop);

		public void AttachBottom(IApplication abstractApp) => InvokeWithNative(abstractApp,AttachBottom);

		// 左吸附：将第二个形状的左边与第一个形状的右边对齐，只移动第二个形状且只水平移动
		public void AttachLeft(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				var baseShape = shapes[1];   // 第一个选中的形状为基准
				var moveShape = shapes[2];   // 第二个选中的形状为要移动的对象

				// 只移动第二个形状的Left属性，使其左边与第一个形状的右边对齐
				moveShape.Left=baseShape.Left+baseShape.Width;

				Toast.Show(ResourceManager.GetString("Align_AttachLeft_Success","已将第二个对象左边与第一个对象右边对齐"),Toast.ToastType.Success);
			});
		}

		public void GuideAlignTop(IApplication abstractApp) => InvokeWithNative(abstractApp,GuideAlignTop);

		public void AttachLeft(IApplication abstractApp) => InvokeWithNative(abstractApp,AttachLeft);

		// 右吸附：将第二个形状的右边与第一个形状的左边对齐，只移动第二个形状且只水平移动
		public void AttachRight(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				var baseShape = shapes[1];   // 第一个选中的形状为基准
				var moveShape = shapes[2];   // 第二个选中的形状为要移动的对象

				// 只移动第二个形状的Left属性，使其右边与第一个形状的左边对齐
				moveShape.Left=baseShape.Left-moveShape.Width;

				Toast.Show(ResourceManager.GetString("Align_AttachRight_Success","已将第二个对象右边与第一个对象左边对齐"),Toast.ToastType.Success);
			});
		}

		public void GuideAlignVCenter(IApplication abstractApp) => InvokeWithNative(abstractApp,GuideAlignVCenter);

		public void AttachRight(IApplication abstractApp) => InvokeWithNative(abstractApp,AttachRight);

		// 上吸附：将第二个形状的上边与第一个形状的下边对齐，只移动第二个形状且只垂直移动
		public void AttachTop(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				var baseShape = shapes[1];   // 第一个选中的形状为基准
				var moveShape = shapes[2];   // 第二个选中的形状为要移动的对象

				// 只移动第二个形状的Top属性，使其上边与第一个形状的下边对齐
				moveShape.Top=baseShape.Top+baseShape.Height;

				Toast.Show(ResourceManager.GetString("Align_AttachTop_Success","已将第二个对象上边与第一个对象下边对齐"),Toast.ToastType.Success);
			});
		}

		public void AttachTop(IApplication abstractApp) => InvokeWithNative(abstractApp,AttachTop);

		// 底对齐到下方最近的水平参考线
		public void GuideAlignBottom(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
			if(sel==null) return;

			ExHandler.Run(() =>
			{
				var guides = netApp.ActivePresentation.Guides;
				List < float > horizontalGuides =[];
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation==NETOP.Enums.PpGuideOrientation.ppHorizontalGuide)
						horizontalGuides.Add(guide.Position);
				}
				if(horizontalGuides.Count==0)
				{
					Toast.Show(ResourceManager.GetString("Align_NoHorizontalGuides","当前文档没有水平参考线"),Toast.ToastType.Warning);
					return;
				}

				List<NETOP.Shape> shapesToProcess = [];

				// 处理不同类型的选择
				if(sel is NETOP.ShapeRange shapeRange)
				{
					shapesToProcess= [.. shapeRange.Cast<NETOP.Shape>()];
				} else if(sel is NETOP.Shape singleShape)
				{
					shapesToProcess= [singleShape];
				} else
				{
					Toast.Show(ResourceManager.GetString("Align_InvalidSelection","无法识别选择的对象类型"),Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in shapesToProcess)
				{
					float bottom = shape.Top + shape.Height;
					// 只考虑在下方的参考线
					var bottomGuides = horizontalGuides.Where(g => g >= bottom).ToList();
					if(bottomGuides.Count==0) continue; // 没有下方参考线则跳过
					float nearest = bottomGuides[0];
					float minDist = Math.Abs(bottom - nearest);
					foreach(var guideY in bottomGuides)
					{
						float dist = Math.Abs(bottom - guideY);
						if(dist<minDist)
						{
							minDist=dist;
							nearest=guideY;
						}
					}
					shape.Top=nearest-shape.Height;
				}
				Toast.Show(ResourceManager.GetString("Align_GuideBottom_Success","已底对齐到参考线"),Toast.ToastType.Success);
			});
		}

		// 水平居中到最近的两条垂直参考线的中点
		public void GuideAlignHCenter(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
			if(sel==null) return;

			ExHandler.Run(() =>
			{
				var guides = netApp.ActivePresentation.Guides;
				List < float > verticalGuides =[];
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation==NETOP.Enums.PpGuideOrientation.ppVerticalGuide)
						verticalGuides.Add(guide.Position);
				}
				if(verticalGuides.Count<2)
				{
					Toast.Show(ResourceManager.GetString("Align_NeedTwoVerticalGuides","至少需要两条垂直参考线"),Toast.ToastType.Warning);
					return;
				}
				verticalGuides.Sort();

				// 统一处理单个形状和形状范围
				IEnumerable<NETOP.Shape> shapesToProcess;
				if(sel is NETOP.Shape singleShape)
				{
					shapesToProcess= [singleShape];
				} else if(sel is NETOP.ShapeRange shapeRange)
				{
					shapesToProcess=shapeRange.Cast<NETOP.Shape>();
				} else
				{
					Toast.Show(ResourceManager.GetString("Align_InvalidSelection","无法识别选中的对象类型"),Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in shapesToProcess)
				{
					float center = shape.Left + (shape.Width / 2f);
					// 找到左侧最近的参考线a和右侧最近的参考线b
					float? a = null, b = null;
					foreach(var g in verticalGuides)
					{
						if(g<=center) a=g;
						if(g>center)
						{
							b=g;
							break;
						}
					}
					if(a==null||b==null) continue; // 没有包围中点的两条参考线则跳过

					float targetCenter = ((float) a + (float) b) / 2f;
					shape.Left=targetCenter-(shape.Width/2f);
				}
				Toast.Show(ResourceManager.GetString("Align_GuideHCenter_Success","已水平居中到参考线"),Toast.ToastType.Success);
			});
		}

		// 左对齐到左侧最近的垂直参考线
		public void GuideAlignLeft(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
			if(sel==null) return;

			ExHandler.Run(() =>
			{
				var guides = netApp.ActivePresentation.Guides;
				List < float > verticalGuides =[];
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation==NETOP.Enums.PpGuideOrientation.ppVerticalGuide)
						verticalGuides.Add(guide.Position);
				}
				if(verticalGuides.Count==0)
				{
					Toast.Show(ResourceManager.GetString("Align_NoVerticalGuides","当前文档没有垂直参考线"),Toast.ToastType.Warning);
					return;
				}

				List<NETOP.Shape> shapesToProcess = [];

				// 处理不同类型的选择
				if(sel is NETOP.ShapeRange shapeRange)
				{
					shapesToProcess= [.. shapeRange.Cast<NETOP.Shape>()];
				} else if(sel is NETOP.Shape singleShape)
				{
					shapesToProcess= [singleShape];
				} else
				{
					Toast.Show(ResourceManager.GetString("Align_InvalidSelection","无法识别选择的对象类型"),Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in shapesToProcess)
				{
					float left = shape.Left;
					// 只考虑在左侧的参考线
					var leftGuides = verticalGuides.Where(g => g <= left).ToList();
					if(leftGuides.Count==0) continue; // 没有左侧参考线则跳过
					float nearest = leftGuides[0];
					float minDist = Math.Abs(left - nearest);
					foreach(var guideX in leftGuides)
					{
						float dist = Math.Abs(left - guideX);
						if(dist<minDist)
						{
							minDist=dist;
							nearest=guideX;
						}
					}
					shape.Left=nearest;
				}
				Toast.Show(ResourceManager.GetString("Align_GuideLeft_Success","已左对齐到参考线"),Toast.ToastType.Success);
			});
		}

		// 右对齐到右侧最近的垂直参考线
		public void GuideAlignRight(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
			if(sel==null) return;

			ExHandler.Run(() =>
			{
				var guides = netApp.ActivePresentation.Guides;
				List<float> verticalGuides = [];
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation==NETOP.Enums.PpGuideOrientation.ppVerticalGuide)
						verticalGuides.Add(guide.Position);
				}
				if(verticalGuides.Count==0)
				{
					Toast.Show(ResourceManager.GetString("Align_NoVerticalGuides","当前文档没有垂直参考线"),Toast.ToastType.Warning);
					return;
				}

				List<NETOP.Shape> shapesToProcess = [];

				// 处理不同类型的选择
				if(sel is NETOP.ShapeRange shapeRange)
				{
					shapesToProcess= [.. shapeRange.Cast<NETOP.Shape>()];
				} else if(sel is NETOP.Shape singleShape)
				{
					shapesToProcess= [singleShape];
				} else
				{
					Toast.Show(ResourceManager.GetString("Align_InvalidSelection","无法识别选择的对象类型"),Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in shapesToProcess)
				{
					float right = shape.Left + shape.Width;
					// 只考虑在右侧的参考线
					var rightGuides = verticalGuides.Where(g => g >= right).ToList();
					if(rightGuides.Count==0) continue; // 没有右侧参考线则跳过
					float nearest = rightGuides[0];
					float minDist = Math.Abs(right - nearest);
					foreach(var guideX in rightGuides)
					{
						float dist = Math.Abs(right - guideX);
						if(dist<minDist)
						{
							minDist=dist;
							nearest=guideX;
						}
					}
					shape.Left=nearest-shape.Width;
				}
				Toast.Show(ResourceManager.GetString("Align_GuideRight_Success","已右对齐到参考线"),Toast.ToastType.Success);
			});
		}

		// 顶对齐到上方最近的水平参考线
		public void GuideAlignTop(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
			if(sel==null) return;

			ExHandler.Run(() =>
			{
				var guides = netApp.ActivePresentation.Guides;
				List<float> horizontalGuides = [];
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation==NETOP.Enums.PpGuideOrientation.ppHorizontalGuide)
						horizontalGuides.Add(guide.Position);
				}
				if(horizontalGuides.Count==0)
				{
					Toast.Show(ResourceManager.GetString("Align_NoHorizontalGuides","当前文档没有水平参考线"),Toast.ToastType.Warning);
					return;
				}

				List<NETOP.Shape> shapesToProcess = [];

				// 处理不同类型的选择
				if(sel is NETOP.ShapeRange shapeRange)
				{
					shapesToProcess= [.. shapeRange.Cast<NETOP.Shape>()];
				} else if(sel is NETOP.Shape singleShape)
				{
					shapesToProcess= [singleShape];
				} else
				{
					Toast.Show(ResourceManager.GetString("Align_InvalidSelection","无法识别选择的对象类型"),Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in shapesToProcess)
				{
					float top = shape.Top;
					// 只考虑在上方的参考线
					var topGuides = horizontalGuides.Where(g => g <= top).ToList();
					if(topGuides.Count==0) continue; // 没有上方参考线则跳过
					float nearest = topGuides[0];
					float minDist = Math.Abs(top - nearest);
					foreach(var guideY in topGuides)
					{
						float dist = Math.Abs(top - guideY);
						if(dist<minDist)
						{
							minDist=dist;
							nearest=guideY;
						}
					}
					shape.Top=nearest;
				}
				Toast.Show(ResourceManager.GetString("Align_GuideTop_Success","已顶对齐到参考线"),Toast.ToastType.Success);
			});
		}

		// 垂直居中到最近的两条水平参考线的中点
		public void GuideAlignVCenter(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
			if(sel==null) return;

			ExHandler.Run(() =>
			{
				var guides = netApp.ActivePresentation.Guides;
				List<float> horizontalGuides = [];
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation==NETOP.Enums.PpGuideOrientation.ppHorizontalGuide)
						horizontalGuides.Add(guide.Position);
				}
				if(horizontalGuides.Count<2)
				{
					Toast.Show(ResourceManager.GetString("Align_NeedTwoHorizontalGuides","至少需要两条水平参考线"),Toast.ToastType.Warning);
					return;
				}
				horizontalGuides.Sort();

				// 统一处理单个形状和形状范围
				IEnumerable<NETOP.Shape> shapesToProcess;
				if(sel is NETOP.Shape singleShape)
				{
					shapesToProcess= [singleShape];
				} else if(sel is NETOP.ShapeRange shapeRange)
				{
					shapesToProcess=shapeRange.Cast<NETOP.Shape>();
				} else
				{
					Toast.Show(ResourceManager.GetString("Align_InvalidSelection","无法识别选中的对象类型"),Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in shapesToProcess)
				{
					float center = shape.Top + (shape.Height / 2f);
					// 找到上方最近的参考线a和下方最近的参考线b
					float? a = null, b = null;
					foreach(var g in horizontalGuides)
					{
						if(g<=center) a=g;
						if(g>center)
						{
							b=g;
							break;
						}
					}
					if(a==null||b==null) continue; // 没有包围中点的两条参考线则跳过

					float targetCenter = ((float) a + (float) b) / 2f;
					shape.Top=targetCenter-(shape.Height/2f);
				}
				Toast.Show(ResourceManager.GetString("Align_GuideVCenter_Success","已垂直居中到参考线"),Toast.ToastType.Success);
			});
		}

		// 高拉伸：高度拉伸到最近两条水平参考线之间并居中
		public void GuidesStretchHeight(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
			if(sel==null) return;

			ExHandler.Run(() =>
			{
				var guides = netApp.ActivePresentation.Guides;
				List < float > horizontalGuides =[];
				// 收集所有水平参考线
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation==NETOP.Enums.PpGuideOrientation.ppHorizontalGuide)
						horizontalGuides.Add(guide.Position);
				}

				// 检查参考线数量
				if(horizontalGuides.Count<2)
				{
					Toast.Show(ResourceManager.GetString("Align_NeedTwoHorizontalGuides","至少需要两条水平参考线"),Toast.ToastType.Warning);
					return;
				}

				// 排序参考线位置（从上到下）
				horizontalGuides.Sort();

				List<NETOP.Shape> shapesToProcess = [];

				// 处理不同类型的选择
				if(sel is NETOP.ShapeRange shapeRange)
				{
					shapesToProcess= [.. shapeRange.Cast<NETOP.Shape>()];
				} else if(sel is NETOP.Shape singleShape)
				{
					shapesToProcess= [singleShape];
				} else
				{
					Toast.Show(ResourceManager.GetString("Align_InvalidSelection","无法识别选择的对象类型"),Toast.ToastType.Warning);
					return;
				}

				// 处理每个选中形状
				foreach(NETOP.Shape shape in shapesToProcess)
				{
					// 计算形状垂直中心
					float centerY = shape.Top + (shape.Height / 2f);
					float? topGuide = null, bottomGuide = null;

					// 查找最近的上下参考线
					foreach(var guideY in horizontalGuides)
					{
						if(guideY<=centerY)
							topGuide=guideY;  // 当前参考线在中心上方
						if(guideY>centerY)
						{
							bottomGuide=guideY;  // 找到中心下方的第一条参考线
							break;
						}
					}

					// 应用参考线位置
					if(topGuide!=null&&bottomGuide!=null)
					{
						shape.Top=(float) topGuide;
						shape.Height=(float) bottomGuide-(float) topGuide;
					}
				}

				Toast.Show(ResourceManager.GetString("Align_StretchHeight_Success","已将高度拉伸到参考线"),Toast.ToastType.Success);
			});
		}

		// 宽高都拉伸：宽度和高度都拉伸到最近两条参考线之间并居中
		public void GuidesStretchSize(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
			if(sel==null) return;

			ExHandler.Run(() =>
			{
				var guides = netApp.ActivePresentation.Guides;
				List < float > verticalGuides =[];
				List < float > horizontalGuides =[];

				// 一次性收集所有参考线
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation==NETOP.Enums.PpGuideOrientation.ppVerticalGuide)
						verticalGuides.Add(guide.Position);
					else if(guide.Orientation==NETOP.Enums.PpGuideOrientation.ppHorizontalGuide)
						horizontalGuides.Add(guide.Position);
				}

				// 检查参考线数量
				if(verticalGuides.Count<2||horizontalGuides.Count<2)
				{
					string message = "";
					if(verticalGuides.Count<2)
						message+=ResourceManager.GetString("Align_NeedTwoVerticalGuides","至少需要两条垂直参考线");
					if(horizontalGuides.Count<2)
					{
						if(message!="") message+=ResourceManager.GetString("Align_And","和");
						message+=ResourceManager.GetString("Align_NeedTwoHorizontalGuides","至少需要两条水平参考线");
					}

					Toast.Show(message,Toast.ToastType.Warning);
					return;
				}

				// 排序参考线
				verticalGuides.Sort();
				horizontalGuides.Sort();

				List<NETOP.Shape> shapesToProcess = [];

				// 处理不同类型的选择
				if(sel is NETOP.ShapeRange shapeRange)
				{
					shapesToProcess= [.. shapeRange.Cast<NETOP.Shape>()];
				} else if(sel is NETOP.Shape singleShape)
				{
					shapesToProcess= [singleShape];
				} else
				{
					Toast.Show(ResourceManager.GetString("Align_InvalidSelection","无法识别选择的对象类型"),Toast.ToastType.Warning);
					return;
				}

				// 处理每个选中形状
				foreach(NETOP.Shape shape in shapesToProcess)
				{
					// 处理高度
					float centerY = shape.Top + (shape.Height / 2f);
					float? topGuide = null, bottomGuide = null;
					foreach(var guideY in horizontalGuides)
					{
						if(guideY<=centerY) topGuide=guideY;
						if(guideY>centerY)
						{
							bottomGuide=guideY;
							break;
						}
					}
					if(topGuide!=null&&bottomGuide!=null)
					{
						shape.Top=(float) topGuide;
						shape.Height=(float) bottomGuide-(float) topGuide;
					}

					// 处理宽度
					float centerX = shape.Left + (shape.Width / 2f);
					float? leftGuide = null, rightGuide = null;
					foreach(var guideX in verticalGuides)
					{
						if(guideX<=centerX) leftGuide=guideX;
						if(guideX>centerX)
						{
							rightGuide=guideX;
							break;
						}
					}
					if(leftGuide!=null&&rightGuide!=null)
					{
						shape.Left=(float) leftGuide;
						shape.Width=(float) rightGuide-(float) leftGuide;
					}
				}

				Toast.Show(ResourceManager.GetString("Align_StretchBoth_Success","已将宽度和高度拉伸到参考线"),Toast.ToastType.Success);
			});
		}

		// 宽拉伸：宽度拉伸到最近两条垂直参考线之间并居中
		public void GuidesStretchWidth(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
			if(sel==null) return;

			ExHandler.Run(() =>
			{
				var guides = netApp.ActivePresentation.Guides;
				List<float> verticalGuides = [];
				foreach(NETOP.Guide guide in guides.Cast<NETOP.Guide>())
				{
					if(guide.Orientation==NETOP.Enums.PpGuideOrientation.ppVerticalGuide)
						verticalGuides.Add(guide.Position);
				}
				if(verticalGuides.Count<2)
				{
					Toast.Show(ResourceManager.GetString("Align_NeedTwoVerticalGuides","至少需要两条垂直参考线"),Toast.ToastType.Warning);
					return;
				}
				verticalGuides.Sort();

				List<NETOP.Shape> shapesToProcess = [];

				// 处理不同类型的选择
				if(sel is NETOP.ShapeRange shapeRange)
				{
					shapesToProcess= [.. shapeRange.Cast<NETOP.Shape>()];
				} else if(sel is NETOP.Shape singleShape)
				{
					shapesToProcess= [singleShape];
				} else
				{
					Toast.Show(ResourceManager.GetString("Align_InvalidSelection","无法识别选择的对象类型"),Toast.ToastType.Warning);
					return;
				}

				foreach(NETOP.Shape shape in shapesToProcess)
				{
					float center = shape.Left + (shape.Width / 2f);
					float? a = null, b = null;
					foreach(var g in verticalGuides)
					{
						if(g<=center) a=g;
						if(g>center)
						{
							b=g;
							break;
						}
					}
					if(a==null||b==null) continue;
					shape.Left=(float) a;
					shape.Width=(float) b-(float) a;
				}
				Toast.Show(ResourceManager.GetString("Align_StretchWidth_Success","已将宽度拉伸到参考线"),Toast.ToastType.Success);
			});
		}

		// 设置选中对象等高
		public void SetEqualHeight(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				float sourceHeight = shapes[1].Height;
				for(int i = 1;i<=shapes.Count;i++)
					shapes[i].Height=sourceHeight;
				Toast.Show(ResourceManager.GetString("Align_SetEqualHeight_Success","已设置等高"),Toast.ToastType.Success);
			});
		}

		// 设置选中对象等宽且等高
		public void SetEqualSize(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				float sourceWidth = shapes[1].Width, sourceHeight = shapes[1].Height;
				for(int i = 1;i<=shapes.Count;i++)
				{
					shapes[i].Width=sourceWidth;
					shapes[i].Height=sourceHeight;
				}
				Toast.Show(ResourceManager.GetString("Align_SetEqualSize_Success","已设置等大小"),Toast.ToastType.Success);
			});
		}

		public void SetEqualSize(IApplication abstractApp) => InvokeWithNative(abstractApp,netApp => SetEqualSize(netApp));

		// 设置选中对象等宽
		public void SetEqualWidth(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				float sourceWidth = shapes[1].Width;
				for(int i = 1;i<=shapes.Count;i++)
					shapes[i].Width=sourceWidth;
				Toast.Show(ResourceManager.GetString("Align_SetEqualWidth_Success","已设置等宽"),Toast.ToastType.Success);
			});
		}

		public void SetEqualWidth(IApplication abstractApp) => InvokeWithNative(abstractApp,netApp => SetEqualWidth(netApp));

		// 下延伸：下边对齐最下侧，上边位置保持不变（高度变大，上边不动）
		public void StretchBottom(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				float maxBottom = float.MinValue;
				for(int i = 1;i<=shapes.Count;i++)
				{
					float bottom = shapes[i].Top + shapes[i].Height;
					if(bottom>maxBottom) maxBottom=bottom;
				}

				for(int i = 1;i<=shapes.Count;i++)
				{
					shapes[i].Height=maxBottom-shapes[i].Top;
				}
				Toast.Show(ResourceManager.GetString("Align_StretchBottom_Success","已向下延伸对齐"),Toast.ToastType.Success);
			});
		}

		// 左延伸：左边对齐最左侧，右边位置保持不变（宽度变大，右边不动）
		public void StretchLeft(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				float minLeft = float.MaxValue;
				for(int i = 1;i<=shapes.Count;i++)
					if(shapes[i].Left<minLeft) minLeft=shapes[i].Left;

				for(int i = 1;i<=shapes.Count;i++)
				{
					float right = shapes[i].Left + shapes[i].Width;
					shapes[i].Width=right-minLeft;
					shapes[i].Left=minLeft;
				}
				Toast.Show(ResourceManager.GetString("Align_StretchLeft_Success","已向左延伸对齐"),Toast.ToastType.Success);
			});
		}

		// 右延伸：右边对齐最右侧，左边位置保持不变（宽度变大，左边不动）
		public void StretchRight(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				float maxRight = float.MinValue;
				for(int i = 1;i<=shapes.Count;i++)
				{
					float right = shapes[i].Left + shapes[i].Width;
					if(right>maxRight) maxRight=right;
				}

				for(int i = 1;i<=shapes.Count;i++)
				{
					shapes[i].Width=maxRight-shapes[i].Left;
					// shapes[i].Left 不变
				}
				Toast.Show(ResourceManager.GetString("Align_StretchRight_Success","已向右延伸对齐"),Toast.ToastType.Success);
			});
		}

		// 上延伸：上边对齐最上侧，下边位置保持不变（高度变大，下边不动）
		public void StretchTop(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				float minTop = float.MaxValue;
				for(int i = 1;i<=shapes.Count;i++)
					if(shapes[i].Top<minTop) minTop=shapes[i].Top;

				for(int i = 1;i<=shapes.Count;i++)
				{
					float bottom = shapes[i].Top + shapes[i].Height;
					shapes[i].Height=bottom-minTop;
					shapes[i].Top=minTop;
				}
				Toast.Show(ResourceManager.GetString("Align_StretchTop_Success","已向上延伸对齐"),Toast.ToastType.Success);
			});
		}

		// 交换两个选中对象的位置和大小
		public void SwapSize(NETOP.Application netApp)
		{
			var abstractApp = GetAbstractApp(netApp);
			var shapes = _shapeHelper.ValidateSelection(abstractApp, true) as dynamic;
			if(shapes==null) return;

			ExHandler.Run(() =>
			{
				var shape1 = shapes[1];
				var shape2 = shapes[2];
				// 交换大小
				(shape1.Width, shape2.Width)=(shape2.Width, shape1.Width);
				(shape1.Height, shape2.Height)=(shape2.Height, shape1.Height);

				// 交换位置
				(shape1.Left, shape2.Left)=(shape2.Left, shape1.Left);
				(shape1.Top, shape2.Top)=(shape2.Top, shape1.Top);

				// 交换填充颜色
				try
				{
					(shape1.Fill.ForeColor.RGB, shape2.Fill.ForeColor.RGB)=(shape2.Fill.ForeColor.RGB, shape1.Fill.ForeColor.RGB);
				} catch { /* 某些形状可能不支持填充颜色 */ }

				// 交换线条样式（安全处理，某些形状可能不支持这些属性）
				try
				{
					if(shape1.Line.Visible==MsoTriState.msoTrue&&shape2.Line.Visible==MsoTriState.msoTrue)
					{
						var dashStyle1 = shape1.Line.DashStyle;
						var dashStyle2 = shape2.Line.DashStyle;
						shape1.Line.DashStyle=dashStyle2;
						shape2.Line.DashStyle=dashStyle1;
					}
				} catch { /* 某些形状可能不支持 DashStyle */ }

				try
				{
					if(shape1.Line.Visible==MsoTriState.msoTrue&&shape2.Line.Visible==MsoTriState.msoTrue)
					{
						var style1 = shape1.Line.Style;
						var style2 = shape2.Line.Style;
						shape1.Line.Style=style2;
						shape2.Line.Style=style1;
					}
				} catch { /* 某些形状可能不支持 Style */ }

				// 交换线条颜色
				try
				{
					if(shape1.Line.Visible==MsoTriState.msoTrue&&shape2.Line.Visible==MsoTriState.msoTrue)
					{
						(shape1.Line.ForeColor.RGB, shape2.Line.ForeColor.RGB)=(shape2.Line.ForeColor.RGB, shape1.Line.ForeColor.RGB);
					}
				} catch { /* 某些形状可能不支持线条颜色 */ }

				// 交换线条宽度
				try
				{
					if(shape1.Line.Visible==MsoTriState.msoTrue&&shape2.Line.Visible==MsoTriState.msoTrue)
					{
						(shape1.Line.Weight, shape2.Line.Weight)=(shape2.Line.Weight, shape1.Line.Weight);
					}
				} catch { /* 某些形状可能不支持线条宽度 */ }

				// 交换透明度
				try
				{
					(shape1.Fill.Transparency, shape2.Fill.Transparency)=(shape2.Fill.Transparency, shape1.Fill.Transparency);
				} catch { /* 某些形状可能不支持透明度 */ }

				// 交换字体字号和颜色
				if(shape1.TextFrame.HasText==MsoTriState.msoTrue&&shape2.TextFrame.HasText==MsoTriState.msoTrue)
				{
					try
					{
						var textRange1 = shape1.TextFrame.TextRange;
						var textRange2 = shape2.TextFrame.TextRange;

						// 交换字体字号
						(textRange1.Font.Size, textRange2.Font.Size)=(textRange2.Font.Size, textRange1.Font.Size);

						// 交换字体颜色
						(textRange1.Font.Color.RGB, textRange2.Font.Color.RGB)=(textRange2.Font.Color.RGB, textRange1.Font.Color.RGB);
					} catch { /* 某些形状可能不支持字体属性 */ }
				}
				Toast.Show(ResourceManager.GetString("Align_SwapSize_Success","已交换大小和位置"),Toast.ToastType.Success);
			});
		}

		/// <summary>
		/// 执行对齐操作（NetOffice 版本）
		/// </summary>
		/// <param name="netApp"> NetOffice PowerPoint 应用程序实例 </param>
		/// <param name="alignment"> 对齐类型 </param>
		/// <param name="alignToSlideMode"> 是否对齐到幻灯片（true：对齐到幻灯片，false：对齐到形状） </param>
		public void ExecuteAlignment(NETOP.Application netApp,AlignmentType alignment,bool alignToSlideMode)
		{
			UndoHelper.BeginUndoEntry(netApp,UndoHelper.UndoNames.AlignShapes);
			ExHandler.Run(() =>
			{
				var abstractApp = GetAbstractApp(netApp);
				var sel = _shapeHelper.ValidateSelection(abstractApp) as dynamic;
				if(sel==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSelection"),Toast.ToastType.Warning);
					return;
				}

				NETOP.ShapeRange shapes;
				// 尝试直接转换为 ShapeRange
				if(sel is NETOP.ShapeRange shapeRange)
				{
					shapes=shapeRange;
				}
				// 如果不是，则尝试处理单个 Shape
				else if(sel is NETOP.Shape shape&&shape.Parent is NETOP.Slide parentSlide)
				{
					// 使用模式匹配确保 Parent 是 Slide 类型，然后创建 ShapeRange
					shapes=parentSlide.Shapes.Range(new object[] { shape.Name });
				} else
				{
					// 如果两种情况都不满足，则选择无效，直接返回
					Toast.Show(ResourceManager.GetString("Toast_InvalidSelection"),Toast.ToastType.Warning);
					return;
				}

				// 判断对齐基准，1.单选形状：总是对齐到幻灯片；2.多选形状：根据按钮状态决定
				MsoTriState alignToSlide = (shapes.Count == 1 || alignToSlideMode) ? MsoTriState.msoTrue : MsoTriState.msoFalse;

				bool TryExecuteMso(string commandName)
				{
					if(string.IsNullOrWhiteSpace(commandName)||_commandExecutor==null)
					{
						return false;
					}

					_logger.LogDebug($"尝试 MSO 命令 '{commandName}'");
					var success = _commandExecutor.ExecuteMso(commandName);
					if(success)
					{
						_logger.LogInformation($"MSO 命令 '{commandName}' 执行成功");
					} else
					{
						_logger.LogDebug($"MSO 命令 '{commandName}' 执行失败");
					}
					return success;
				}

				// 注意：对齐基准已经在切换按钮（Tb101）点击时通过 MSO 命令设置 ObjectsAlignRelativeToContainerSmart 或
				// ObjectsAlignSelectedSmart 因此这里不需要再次设置基准

				// 执行对齐操作
				switch(alignment)
				{
					case AlignmentType.Left:
						// _commandExecutor.ExecuteMenuPath("文件|另存为"); _commandExecutor.ExecuteCommandById(748);
						if(TryExecuteMso(ObjectsAlignLeftSmart))
						{
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						} else
						{
							shapes.Align(MsoAlignCmd.msoAlignLefts,alignToSlide);
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						}
						break;

					case AlignmentType.Right:
						if(TryExecuteMso(ObjectsAlignRightSmart))
						{
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						} else
						{
							shapes.Align(MsoAlignCmd.msoAlignRights,alignToSlide);
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						}
						break;

					case AlignmentType.Top:
						if(TryExecuteMso(ObjectsAlignTopSmart))
						{
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						} else
						{
							shapes.Align(MsoAlignCmd.msoAlignTops,alignToSlide);
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						}
						break;

					case AlignmentType.Bottom:
						if(TryExecuteMso(ObjectsAlignBottomSmart))
						{
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						} else
						{
							shapes.Align(MsoAlignCmd.msoAlignBottoms,alignToSlide);
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						}
						break;

					case AlignmentType.Centers:
						if(TryExecuteMso(ObjectsAlignCenterHorizontalSmart))
						{
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						} else
						{
							shapes.Align(MsoAlignCmd.msoAlignCenters,alignToSlide);
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						}
						break;

					case AlignmentType.Middles:
						if(TryExecuteMso(ObjectsAlignMiddleVerticalSmart))
						{
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						} else
						{
							shapes.Align(MsoAlignCmd.msoAlignMiddles,alignToSlide);
							Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
						}
						break;

					case AlignmentType.Horizontally:
					{
						// 根据对齐基准确定所需的最小形状数
						int minRequired = (alignToSlide == MsoTriState.msoTrue) ? 1 : 3;
						if(shapes.Count>=minRequired)
						{
							if(TryExecuteMso(AlignDistributeHorizontally))
							{
								Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
							} else
							{
								shapes.Distribute(MsoDistributeCmd.msoDistributeHorizontally,alignToSlide);
								Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
							}
						} else
						{
							string basis = (alignToSlide == MsoTriState.msoTrue)
									? ResourceManager.GetString("Toast_Basis_Page", "页面")
									: ResourceManager.GetString("Toast_Basis_Shape", "形状");
							Toast.Show(ResourceManager.GetString("Toast_AlignMinShapes",basis,minRequired),Toast.ToastType.Warning);
						}
					}
					break;

					case AlignmentType.Vertically:
					{
						// 根据对齐基准确定所需的最小形状数
						int minRequired = (alignToSlide == MsoTriState.msoTrue) ? 1 : 3;
						if(shapes.Count>=minRequired)
						{
							if(TryExecuteMso(AlignDistributeVertically))
							{
								Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
							} else
							{
								shapes.Distribute(MsoDistributeCmd.msoDistributeVertically,alignToSlide);
								Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"),Toast.ToastType.Success);
							}
						} else
						{
							string basis = (alignToSlide == MsoTriState.msoTrue)
									? ResourceManager.GetString("Toast_Basis_Page", "页面")
									: ResourceManager.GetString("Toast_Basis_Shape", "形状");
							Toast.Show(ResourceManager.GetString("Toast_AlignMinShapes",basis,minRequired),Toast.ToastType.Warning);
						}
					}
					break;

					default:
						Toast.Show(ResourceManager.GetString("Toast_UnknownAlignment",alignment.ToString()),Toast.ToastType.Error);
						break;
				}
			});
		}

		/// <summary>
		/// 执行对齐操作（抽象接口版本）
		/// </summary>
		/// <param name="abstractApp"> 抽象应用程序实例 </param>
		/// <param name="alignment"> 对齐类型 </param>
		/// <param name="alignToSlideMode"> 是否对齐到幻灯片 </param>
		public void ExecuteAlignment(IApplication abstractApp,AlignmentType alignment,bool alignToSlideMode)
		{
			// 使用统一的 ApplicationHelper 获取 NetOffice Application 对象
			var netApp = ApplicationHelper.GetNetOfficeApplication(abstractApp);
			if(netApp==null)
			{
				_logger.LogError("无法获取 NetOffice Application 对象");
				return;
			}

			ExecuteAlignment(netApp,alignment,alignToSlideMode);
		}

		#endregion Public Methods
	}
}

using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Logging;
using PPA.Utilities;
using System;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Manipulation
{
	/// <summary>
	/// 图表批量操作辅助类
	/// </summary>
	internal class ChartBatchHelper(IChartFormatHelper chartFormatHelper,IShapeHelper shapeHelper,ILogger logger = null):IChartBatchHelper
	{
		private readonly IChartFormatHelper _chartFormatHelper = chartFormatHelper??throw new ArgumentNullException(nameof(chartFormatHelper));
		private readonly IShapeHelper _shapeHelper = shapeHelper??throw new ArgumentNullException(nameof(shapeHelper));
		private readonly ILogger _logger = logger??LoggerProvider.GetLogger();

		#region IChartBatchHelper 实现

		public void FormatCharts(NETOP.Application netApp)
		{
			if(netApp==null) throw new ArgumentNullException(nameof(netApp));
			FormatChartsInternal(netApp,_chartFormatHelper);
		}

		#endregion IChartBatchHelper 实现

		#region 内部实现

		private void FormatChartsInternal(NETOP.Application netApp,IChartFormatHelper chartFormatHelper)
		{
			_logger.LogInformation($"启动，netApp类型={netApp?.GetType().Name??"null"}");
			if(chartFormatHelper==null)
				throw new InvalidOperationException("无法获取 IChartFormatHelper 服务");

			ExHandler.Run(() =>
			{
				var currentApp = netApp;
				currentApp=ApplicationHelper.EnsureValidNetApplication(currentApp);
				if(currentApp==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_NoSlide"),Toast.ToastType.Warning);
					return;
				}

				var selection = GetSelectionWithRetry(ref currentApp);

				// 调试：记录选中对象信息
				if(selection==null)
				{
					_logger.LogWarning("ValidateSelection 返回 null，没有选中对象");
					Toast.Show(ResourceManager.GetString("Toast_FormatCharts_NoSelection"),Toast.ToastType.Warning);
					return;
				}

				UndoHelper.BeginUndoEntry(currentApp,UndoHelper.UndoNames.FormatCharts);

				// 调试：记录选中对象的数量和类型
				try
				{
					if(selection is NETOP.ShapeRange shapeRange)
					{
						int count = ExHandler.SafeGet(() => shapeRange.Count, defaultValue: 0);
						_logger.LogInformation($"选中对象类型=ShapeRange, 数量={count}");
					} else if(selection is NETOP.Shape shape)
					{
						_logger.LogInformation("选中对象类型=Shape, 数量=1");
					} else
					{
						_logger.LogInformation($"选中对象类型={selection?.GetType().Name??"null"}");
					}
				} catch(Exception ex)
				{
					_logger.LogWarning($"获取选中对象信息失败: {ex.Message}");
				}

				// 直接处理选中的图表形状，避免 COM 对象生命周期问题（与美化文本保持一致）
				ProcessChartsFromSelection(selection,currentApp,chartFormatHelper);
			},enableTiming: true);
		}

		/// <summary>
		/// 检查形状是否是图表
		/// </summary>
		private bool IsChartShape(NETOP.Shape shape)
		{
			if(shape==null) return false;

			// 1. 先检查形状类型 msoChart (3) 是普通图表 msoPlaceholder (14) 是占位符，可能包含图表
			var shapeType = ExHandler.SafeGet(() => shape.Type, defaultValue: MsoShapeType.msoAutoShape);

			// 如果既不是 Chart 也不是 Placeholder，那通常不可能是图表 注意：如果未来发现其他类型（如 Group）也可能包含图表，需要在这里添加
			if(shapeType!=MsoShapeType.msoChart&&shapeType!=MsoShapeType.msoPlaceholder)
			{
				return false;
			}

			// 2. 再次确认 HasChart 属性 (双重检查)
			return ExHandler.SafeGet(() => shape.HasChart==MsoTriState.msoTrue,defaultValue: false);
		}

		/// <summary>
		/// 从选区处理图表形状（与美化文本的处理方式保持一致，避免 COM 对象生命周期问题）
		/// </summary>
		private void ProcessChartsFromSelection(object selection,NETOP.Application netApp,IChartFormatHelper chartFormatHelper)
		{
			var shapeSelection = ShapeSelectionFactory.Create(selection);
			if(shapeSelection==null)
			{
				Toast.Show(ResourceManager.GetString("Toast_FormatCharts_NoSelection"),Toast.ToastType.Warning);
				return;
			}

			int count = 0;
			foreach(var shape in shapeSelection)
			{
				if(shape!=null&&IsChartShape(shape))
				{
					chartFormatHelper.FormatChartText(shape);
					count++;
				}
			}

			if(count>0)
			{
				Toast.Show(ResourceManager.GetString("Toast_FormatCharts_Success",count),Toast.ToastType.Success);
			} else
			{
				Toast.Show(ResourceManager.GetString("Toast_FormatCharts_NoSelection"),Toast.ToastType.Warning);
			}
		}

		private dynamic GetSelectionWithRetry(ref NETOP.Application netApp)
		{
			var selection = _shapeHelper.ValidateSelection(netApp, showWarningWhenInvalid: false);
			if(selection!=null)
			{
				return selection;
			}

			_logger.LogWarning("返回 null，尝试刷新 Application 后重试");
			netApp=ApplicationHelper.EnsureValidNetApplication(netApp);
			if(netApp==null)
			{
				return null;
			}

			return _shapeHelper.ValidateSelection(netApp,showWarningWhenInvalid: false);
		}

		#endregion 内部实现
	}
}

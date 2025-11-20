using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using PPA.Core.Logging;
using System;
using System.Collections.Generic;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 图表格式化辅助类 提供图表的格式化功能
	/// </summary>
	/// <remarks> 构造函数，通过依赖注入获取配置和服务 </remarks>
	/// <param name="config"> 格式化配置 </param>
	/// <param name="shapeHelper"> 形状工具服务（可选） </param>
	internal class ChartFormatHelper(IFormattingConfig config,ILogger logger = null):IChartFormatHelper
	{
		private readonly IFormattingConfig _config = config ?? throw new ArgumentNullException(nameof(config));
		private readonly ILogger _logger = logger ?? LoggerProvider.GetLogger();

		/// <summary>
		/// 格式化图表文本
		/// </summary>
		/// <param name="shape"> 包含图表的形状对象 </param>
		public void FormatChartText(NETOP.Shape shape)
		{
			if(shape==null)
			{
				_logger.LogWarning("shape 为 null，返回");
				return;
			}

			// 使用与 IsChartShape 完全相同的检查逻辑：先检查 HasChart，如果失败再尝试获取 Chart 某些情况下 HasChart 可能不准确，所以即使
			// HasChart 为 false，也尝试直接获取 Chart
			bool hasChart = ExHandler.SafeGet(() => shape.HasChart == MsoTriState.msoTrue, defaultValue: false);
			NETOP.Chart chart = ExHandler.SafeGet(() => shape.Chart, defaultValue: (NETOP.Chart)null);

			// 如果 HasChart 为 true 但 Chart 为 null，记录警告
			if(hasChart&&chart==null)
			{
				_logger.LogWarning("HasChart 为 true 但无法获取 Chart 对象");
			}

			if(chart==null)
			{
				_logger.LogWarning("无法获取 Chart 对象");
				return;
			}

			// 从配置加载参数
			var config = _config.Chart;
			string fontFamily = config.RegularFont.Name;
			float regularSize = config.RegularFont.Size;
			float titleSize = config.TitleFont.Size;
			bool titleBold = config.TitleFont.Bold;

			// 设置图表各部分的字体
			SetChartTitleFont(chart,fontFamily,titleSize,titleBold);
			SetChartLegendFont(chart,fontFamily,regularSize);
			SetChartDataTableFont(chart,fontFamily,regularSize);

			SetChartDataLabelsFont(chart,fontFamily,regularSize);
			SetChartAxesFont(chart,regularSize);
		}

		/// <summary>
		/// 格式化图表文本（抽象接口版本）
		/// </summary>
		/// <param name="shape"> 包含图表的抽象形状对象 </param>
		public void FormatChartText(IShape shape)
		{
			if(shape==null)
			{
				_logger.LogWarning("shape 为 null，返回");
				return;
			}

			// 使用 AdapterUtils 统一转换
			var native = AdapterUtils.UnwrapShape(shape);
			if(native!=null)
			{
				_logger.LogInformation("成功获取 NetOffice.Shape，使用 PowerPoint 格式化");
				FormatChartText(native);
				return;
			}

			_logger.LogError($"未知的形状类型 {shape.GetType().FullName}，无法格式化");
		}
		/// <summary>
		/// 设置图表标题字体
		/// </summary>
		private void SetChartTitleFont(NETOP.Chart chart,string fontFamily,float size,bool bold)
		{
			ExHandler.Run(() =>
			{
				bool hasTitle = ExHandler.SafeGet(() => chart.HasTitle, defaultValue: false);
				if(!hasTitle)
				{
					return;
				}

				var chartTitle = ExHandler.SafeGet(() => chart.ChartTitle, defaultValue: (NETOP.ChartTitle)null);
				if(chartTitle==null)
				{
					_logger.LogWarning("ChartTitle 为空，无法设置标题字体");
					return;
				}

				var font = ExHandler.SafeGet(() => chartTitle.Font, defaultValue: (NETOP.ChartFont)null);
				if(font==null)
				{
					_logger.LogWarning("ChartTitle.Font 为空，无法设置标题字体");
					return;
				}

				font.Name=fontFamily;
				font.Bold=bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
				font.Size=size;
			},message: $"{nameof(ChartFormatHelper)}.{nameof(SetChartTitleFont)}");
		}

		/// <summary>
		/// 设置图例字体
		/// </summary>
		private void SetChartLegendFont(NETOP.Chart chart,string fontFamily,float size)
		{
			ExHandler.Run(() =>
			{
				bool hasLegend = ExHandler.SafeGet(() => chart.HasLegend, defaultValue: false);
				if(!hasLegend)
				{
					return;
				}

				var legend = ExHandler.SafeGet(() => chart.Legend, defaultValue: (NETOP.Legend)null);
				if(legend==null)
				{
					_logger.LogWarning("Legend 为空，无法设置图例字体");
					return;
				}

				var font = ExHandler.SafeGet(() => legend.Font, defaultValue: (NETOP.ChartFont)null);
				if(font==null)
				{
					_logger.LogWarning("Legend.Font 为空，无法设置图例字体");
					return;
				}

				font.Name=fontFamily;
				font.Size=size;
			},message: $"{nameof(ChartFormatHelper)}.{nameof(SetChartLegendFont)}");
		}

		/// <summary>
		/// 设置数据表字体
		/// </summary>
		private void SetChartDataTableFont(NETOP.Chart chart,string fontFamily,float size)
		{
			ExHandler.Run(() =>
			{
				bool hasDataTable = ExHandler.SafeGet(() => chart.HasDataTable, defaultValue: false);
				if(!hasDataTable)
				{
					return;
				}

				var dataTable = ExHandler.SafeGet(() => chart.DataTable, defaultValue: (object)null);
				if(dataTable==null)
				{
					_logger.LogWarning("DataTable 为空，无法设置数据表字体");
					return;
				}

				var font = ExHandler.SafeGet(() => ((dynamic)dataTable).Font, defaultValue: (NETOP.ChartFont)null);
				if(font==null)
				{
					_logger.LogWarning("DataTable.Font 为空，无法设置数据表字体");
					return;
				}

				font.Name=fontFamily;
				font.Size=size;
			},message: $"{nameof(ChartFormatHelper)}.{nameof(SetChartDataTableFont)}");
		}

		/// <summary>
		/// 设置图表中所有系列的数据标签字体。
		/// </summary>
		/// <param name="chart"> 目标图表对象。 </param>
		/// <param name="fontFamily"> 字体名称。 </param>
		/// <param name="size"> 字体大小。 </param>
		private void SetChartDataLabelsFont(NETOP.Chart chart,string fontFamily,float size)
		{
			// 1. 获取强类型的SeriesCollection，避免使用dynamic
			var seriesCollection = ExHandler.SafeGet(() => chart.SeriesCollection() as NETOP.SeriesCollection, defaultValue: (NETOP.SeriesCollection)null);
			if(seriesCollection==null)
			{
				_logger.LogWarning("无法获取图表的SeriesCollection。");
				return;
			}

			// 2. 使用 for 循环，这是遍历COM集合最可靠的方式 COM集合索引通常从1开始
			int seriesCount = ExHandler.SafeGet(() => seriesCollection.Count, defaultValue: 0);
			for(int i = 1;i<=seriesCount;i++)
			{
				try
				{
					// 3. 通过索引获取强类型的Series对象
					var series = ExHandler.SafeGet(() => seriesCollection[i], defaultValue: (NETOP.Series)null);
					if(series==null) continue;

					// 调用辅助方法处理单个系列
					SetDataLabelsFontForSeries(series,fontFamily,size);
				} catch(Exception ex)
				{
					_logger.LogError($"处理系列 {i} 时出错: {ex.Message}",ex);
					// 继续处理下一个系列
				}
			}
		}

		/// <summary>
		/// 为单个系列设置数据标签字体。
		/// </summary>
		private void SetDataLabelsFontForSeries(NETOP.Series series,string fontFamily,float size)
		{
			// 检查系列是否有数据标签
			bool hasDataLabels = ExHandler.SafeGet(() => series.HasDataLabels, defaultValue: false);
			if(!hasDataLabels) return;

			// DataLabels 是一个方法，需要调用它
			var dataLabels = ExHandler.SafeGet(() => series.DataLabels() as NETOP.DataLabels, defaultValue: (NETOP.DataLabels)null);
			if(dataLabels==null) return;

			// 获取字体对象
			var font = ExHandler.SafeGet(() => dataLabels.Font, defaultValue: (NETOP.ChartFont)null);
			if(font!=null)
			{
				font.Name=fontFamily;
				font.Size=size;
			}
		}

		/// <summary>
		/// 设置坐标轴字体
		/// </summary>
		private void SetChartAxesFont(NETOP.Chart chart,float size)
		{
			ExHandler.Run(() =>
			{
				var chartType = ExHandler.SafeGet(() => chart.ChartType, defaultValue: XlChartType.xlColumnClustered);
				var nonAxisCharts = new HashSet<XlChartType>
				{
					XlChartType.xlPie, XlChartType.xl3DPie, XlChartType.xlDoughnut,
					XlChartType.xlPieOfPie, XlChartType.xlBarOfPie,
					XlChartType.xlRadar, XlChartType.xlRadarFilled
				};

				if(nonAxisCharts.Contains(chartType))
					return;

				SafeSetAxis(chart,XlAxisType.xlCategory,XlAxisGroup.xlPrimary,size);
				SafeSetAxis(chart,XlAxisType.xlValue,XlAxisGroup.xlPrimary,size);
				SafeSetAxis(chart,XlAxisType.xlCategory,XlAxisGroup.xlSecondary,size);
				SafeSetAxis(chart,XlAxisType.xlValue,XlAxisGroup.xlSecondary,size);
			});
		}

		/// <summary>
		/// 安全地设置图表坐标轴的字体。
		/// </summary>
		/// <param name="chart"> 目标图表对象。 </param>
		/// <param name="axisType"> 坐标轴类型（主/次坐标轴）。 </param>
		/// <param name="axisGroup"> 坐标轴组。 </param>
		/// <param name="size"> 字体大小。 </param>
		private void SafeSetAxis(NETOP.Chart chart,XlAxisType axisType,XlAxisGroup axisGroup,float size)
		{
			// 从配置加载字体设置
			var config = _config.Chart;
			string fontFamily = config.RegularFont.Name;

			// 使用 SafeGet 安全地获取坐标轴对象
			NETOP.Axis axis = ExHandler.SafeGet(() => (NETOP.Axis)chart.Axes(axisType, axisGroup), defaultValue: (NETOP.Axis)null);
			if(axis==null)
			{
				// 对于次坐标轴，如果不存在是正常的，只记录 DEBUG 级别日志
				if(axisGroup==XlAxisGroup.xlSecondary)
				{
					_logger.LogDebug($"图表不支持 {axisType}-{axisGroup} 坐标轴（这是正常的）");
				} else
				{
					_logger.LogWarning($"坐标轴 {axisType}-{axisGroup} 对象为 null");
				}
				return;
			}

			// --- 优化核心：提取公共逻辑，减少重复代码 ---
			// 1. 设置刻度线标签字体（使用 SafeGet 安全获取）
			var tickLabels = ExHandler.SafeGet(() => axis.TickLabels, defaultValue: (NETOP.TickLabels)null);
			if(tickLabels!=null)
			{
				TrySetFont(tickLabels,fontFamily,size,"刻度标签");
			}

			// 2. 设置坐标轴标题字体 (仅在标题存在时)
			bool hasTitle = ExHandler.SafeGet(() => axis.HasTitle, defaultValue: false);
			if(hasTitle)
			{
				var axisTitle = ExHandler.SafeGet(() => axis.AxisTitle, defaultValue: (NETOP.AxisTitle)null);
				if(axisTitle!=null)
				{
					TrySetFont(axisTitle,fontFamily,size,"坐标轴标题");
				}
			}
		}

		/// <summary>
		/// 尝试为具有Font属性的图表元素设置字体。
		/// </summary>
		/// <param name="element"> 图表元素（如TickLabels, AxisTitle）。 </param>
		/// <param name="fontFamily"> 字体名称。 </param>
		/// <param name="size"> 字体大小。 </param>
		/// <param name="elementName"> 元素名称，用于日志记录。 </param>
		private void TrySetFont(object element,string fontFamily,float size,string elementName)
		{
			if(element==null)
			{
				_logger.LogDebug($"{elementName} 对象为 null，跳过设置。");
				return;
			}

			try
			{
				var fontProperty = element.GetType().GetProperty("Font");
				var font = fontProperty?.GetValue(element);

				if(font is NETOP.ChartFont chartFont)
				{
					chartFont.Name=fontFamily;
					chartFont.Size=size;
					_logger.LogDebug($"成功设置 {elementName} 字体为 {fontFamily} {size}pt。");
				} else
				{
					_logger.LogWarning($"无法从 {elementName} 获取有效的 ChartFont 对象。");
				}
			} catch(Exception ex)
			{
				_logger.LogError($"设置 {elementName} 字体时出错: {ex.Message}",ex);
			}
		}
	}
}

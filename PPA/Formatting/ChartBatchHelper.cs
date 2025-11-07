using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Shape;
using PPA.Utilities;
using System.Collections.Generic;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 图表批量操作辅助类
	/// </summary>
	public static class ChartBatchHelper
	{
		/// <summary>
		/// 批量格式化图表
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		public static void Bt503_Click(NETOP.Application app)
		{
			UndoHelper.BeginUndoEntry(app,UndoHelper.UndoNames.FormatCharts);

			ExHandler.Run(() =>
			{
				var slide = ShapeUtils.TryGetCurrentSlide(app);
				if(slide==null) return;

				// 收集需要处理的图表形状
				var chartShapes = new List<NETOP.Shape>();
				var sel = ShapeUtils.ValidateSelection(app);

				if(sel!=null)
				{
					// 处理选中的对象
					if(sel is NETOP.Shape shape&&shape.HasChart==MsoTriState.msoTrue)
					{
						chartShapes.Add(shape);
					} else if(sel is NETOP.ShapeRange shapes)
					{
						foreach(NETOP.Shape s in shapes)
						{
							if(s.HasChart==MsoTriState.msoTrue)
								chartShapes.Add(s);
						}
					}
				} else
				{
					// 处理当前幻灯片上所有对象
					foreach(NETOP.Shape shape in slide.Shapes)
					{
						if(shape.HasChart==MsoTriState.msoTrue)
							chartShapes.Add(shape);
					}
				}

				// 格式化所有图表
				foreach(var shape in chartShapes)
				{
					ChartFormatHelper.FormatChartText(shape);
				}

				// 显示结果
				if(chartShapes.Count>0)
					Toast.Show(ResourceManager.GetString("Toast_FormatCharts_Success",chartShapes.Count),Toast.ToastType.Success);
				else
					Toast.Show(sel!=null ? ResourceManager.GetString("Toast_FormatCharts_NoSelection") : ResourceManager.GetString("Toast_FormatCharts_NoCharts"),Toast.ToastType.Info);
			});
		}
	}
}

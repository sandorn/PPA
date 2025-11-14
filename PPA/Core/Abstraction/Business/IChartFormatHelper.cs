using NETOP = NetOffice.PowerPointApi;
using PPA.Core.Abstraction.Presentation;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 图表格式化辅助接口
	/// 注意：当前使用 NetOffice 类型，后续阶段将改为使用平台抽象接口
	/// </summary>
	public interface IChartFormatHelper
	{
		/// <summary>
		/// 格式化图表文本
		/// </summary>
		/// <param name="shape">包含图表的形状对象</param>
		void FormatChartText(NETOP.Shape shape);

		/// <summary>
		/// 格式化图表文本（抽象接口版本）
		/// </summary>
		/// <param name="shape">包含图表的抽象形状对象</param>
		void FormatChartText(IShape shape);
	}
}


using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 图表格式化辅助接口 提供图表的格式化功能，包括标题、图例、数据标签、坐标轴等元素的字体设置
	/// </summary>
	/// <remarks> 此接口定义了图表格式化的接口，通过依赖注入使用，便于测试和扩展。 格式化参数从配置文件（ <see cref="IFormattingConfig" />）中读取。 </remarks>
	public interface IChartFormatHelper
	{
		/// <summary>
		/// 格式化图表文本
		/// </summary>
		/// <param name="shape"> 包含图表的 NetOffice 形状对象，不能为 null </param>
		/// <remarks>
		/// 此方法会格式化图表中的以下文本元素：
		/// <list type="bullet">
		/// <item>
		/// <description> 图表标题字体 </description>
		/// </item>
		/// <item>
		/// <description> 图例字体 </description>
		/// </item>
		/// <item>
		/// <description> 数据标签字体 </description>
		/// </item>
		/// <item>
		/// <description> 坐标轴字体（包括主坐标轴和次坐标轴） </description>
		/// </item>
		/// <item>
		/// <description> 数据表字体（如果存在） </description>
		/// </item>
		/// </list>
		/// 如果图表不支持某些元素（如次坐标轴），会安全地跳过这些元素。
		/// </remarks>
		void FormatChartText(NETOP.Shape shape);
	}
}

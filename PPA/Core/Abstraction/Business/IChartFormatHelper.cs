using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 图表格式化辅助接口 提供图表的格式化功能，包括标题、图例、数据标签、坐标轴等元素的字体设置
	/// </summary>
	/// <remarks>
	/// 此接口定义了图表格式化的接口，通过依赖注入使用，便于测试和扩展。 格式化参数从配置文件（ <see cref="IFormattingConfig" />）中读取。
	/// <para> <strong> 接口版本说明： </strong> </para>
	/// <list type="bullet">
	/// <item>
	/// <description>
	/// <strong> NetOffice 版本 </strong>（ <see cref="FormatChartText(NETOP.Shape)" />）：
	/// 提供完整的图表格式化功能，包括标题、图例、数据标签、坐标轴等所有元素的字体设置。 这是主要使用的版本，功能最完整。
	/// </description>
	/// </item>
	/// <item>
	/// <description>
	/// <strong> 抽象接口版本 </strong>（ <see cref="FormatChartText(IShape)" />）： 内部通过适配器模式转换为 NetOffice
	/// 对象后调用 NetOffice 版本。 主要用于与抽象接口系统的集成，功能覆盖度依赖于抽象接口的实现。
	/// </description>
	/// </item>
	/// </list>
	/// </remarks>
	public interface IChartFormatHelper
	{
		/// <summary>
		/// 格式化图表文本（NetOffice 版本）
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

		/// <summary>
		/// 格式化图表文本（抽象接口版本）
		/// </summary>
		/// <param name="shape"> 包含图表的抽象形状对象，不能为 null </param>
		/// <remarks> 此方法内部会将抽象接口转换为 NetOffice 对象，然后调用 NetOffice 版本的方法。 </remarks>
		void FormatChartText(IShape shape);
	}
}

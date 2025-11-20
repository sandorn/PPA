using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 图表批量操作辅助接口 提供图表的批量格式化功能
	/// </summary>
	/// <remarks>
	/// 此接口定义了图表批量操作的接口，通过依赖注入使用，便于测试和扩展。 注意：当前方法签名使用
	/// <see cref="NETOP.Application" />，后续阶段将改为使用平台抽象接口（ <see cref="IApplication" />）。
	/// </remarks>
	public interface IChartBatchHelper
	{
		/// <summary>
		/// 批量格式化图表 对选中的图表形状或当前幻灯片上的所有图表形状进行格式化
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例，不能为 null </param>
		/// <remarks>
		/// 格式化行为：
		/// <list type="bullet">
		/// <item>
		/// <description> 如果选中了图表形状，则格式化选中的图表 </description>
		/// </item>
		/// <item>
		/// <description> 如果没有选中图表，则格式化当前幻灯片上的所有图表形状 </description>
		/// </item>
		/// </list>
		/// 格式化包括：图表标题、图例、数据标签、坐标轴等元素的字体设置。 格式化参数从配置文件（ <see cref="IFormattingConfig" />）中读取。
		/// </remarks>
		void FormatCharts(NETOP.Application app);
	}
}

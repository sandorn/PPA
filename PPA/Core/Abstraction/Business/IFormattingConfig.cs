using PPA.Manipulation.Config;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 格式化配置接口 用于抽象配置访问，支持依赖注入和单元测试
	/// </summary>
	/// <remarks>
	/// 此接口定义了格式化配置的访问接口，所有格式化参数都通过此接口获取。 配置数据从 XML 配置文件（PPAConfig.xml）中加载，支持运行时修改和持久化。 实现类： <see cref="FormattingConfig" />。
	/// </remarks>
	public interface IFormattingConfig
	{
		/// <summary>
		/// 获取表格格式化配置
		/// </summary>
		/// <value> 表格格式化配置对象，包含表头、数据行、边框等样式设置 </value>
		TableFormattingConfig Table { get; }

		/// <summary>
		/// 获取文本格式化配置
		/// </summary>
		/// <value> 文本格式化配置对象，包含字体、颜色、边距等样式设置 </value>
		TextFormattingConfig Text { get; }

		/// <summary>
		/// 获取图表格式化配置
		/// </summary>
		/// <value> 图表格式化配置对象，包含标题、图例、坐标轴等字体设置 </value>
		ChartFormattingConfig Chart { get; }

		/// <summary>
		/// 获取快捷键配置
		/// </summary>
		/// <value> 快捷键配置对象，包含各种操作的快捷键设置 </value>
		ShortcutsConfig Shortcuts { get; }
	}
}

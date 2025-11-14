using PPA.Formatting;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 格式化配置接口
	/// 用于抽象配置访问，支持依赖注入
	/// </summary>
	public interface IFormattingConfig
	{
		/// <summary>
		/// 表格格式化配置
		/// </summary>
		TableFormattingConfig Table { get; }

		/// <summary>
		/// 文本格式化配置
		/// </summary>
		TextFormattingConfig Text { get; }

		/// <summary>
		/// 图表格式化配置
		/// </summary>
		ChartFormattingConfig Chart { get; }

		/// <summary>
		/// 快捷键配置
		/// </summary>
		ShortcutsConfig Shortcuts { get; }
	}
}


using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 文本批量操作辅助接口 提供文本的批量格式化功能
	/// </summary>
	/// <remarks>
	/// 此接口定义了文本批量操作的接口，通过依赖注入使用，便于测试和扩展。 注意：当前方法签名使用
	/// <see cref="NETOP.Application" />，后续阶段将改为使用平台抽象接口（ <see cref="IApplication" />）。
	/// </remarks>
	public interface ITextBatchHelper
	{
		/// <summary>
		/// 批量格式化文本 对选中的文本形状或当前幻灯片上的所有文本形状进行格式化
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例，不能为 null </param>
		/// <remarks>
		/// 格式化行为：
		/// <list type="bullet">
		/// <item>
		/// <description> 如果选中了文本形状，则格式化选中的文本 </description>
		/// </item>
		/// <item>
		/// <description> 如果没有选中文本，则格式化当前幻灯片上的所有文本形状 </description>
		/// </item>
		/// </list>
		/// 格式化参数从配置文件（ <see cref="IFormattingConfig" />）中读取。
		/// </remarks>
		void FormatText(NETOP.Application app);
	}
}

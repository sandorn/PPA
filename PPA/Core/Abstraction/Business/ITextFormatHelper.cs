using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 文本格式化辅助接口 提供文本形状的格式化功能，包括字体、颜色、边距等样式设置
	/// </summary>
	/// <remarks>
	/// 此接口定义了文本格式化的接口，通过依赖注入使用，便于测试和扩展。 格式化参数从配置文件（ <see cref="IFormattingConfig" />）中读取。
	/// <para> <strong> 接口版本说明： </strong> </para>
	/// <list type="bullet">
	/// <item>
	/// <description>
	/// <strong> NetOffice 版本 </strong>（ <see cref="ApplyTextFormatting(NETOP.Shape)" />）：
	/// 提供完整的文本格式化功能，包括字体、颜色、边距、段落对齐、项目符号等所有功能。 这是主要使用的版本，功能最完整。
	/// </description>
	/// </item>
	/// <item>
	/// <description>
	/// <strong> 抽象接口版本 </strong>（ <see cref="ApplyTextFormatting(IShape)" />）： 内部通过适配器模式转换为
	/// NetOffice 对象后调用 NetOffice 版本。 主要用于与抽象接口系统的集成，功能覆盖度依赖于抽象接口的实现。
	/// </description>
	/// </item>
	/// </list>
	/// </remarks>
	public interface ITextFormatHelper
	{
		/// <summary>
		/// 应用文本格式化到指定形状（NetOffice 版本）
		/// </summary>
		/// <param name="shp"> 要格式化的 NetOffice 形状对象，不能为 null </param>
		/// <remarks>
		/// 此方法会应用以下格式化设置：
		/// <list type="bullet">
		/// <item>
		/// <description> 文本框边距（上、下、左、右） </description>
		/// </item>
		/// <item>
		/// <description> 字体属性（名称、大小、颜色、加粗、斜体等） </description>
		/// </item>
		/// <item>
		/// <description> 段落对齐方式 </description>
		/// </item>
		/// <item>
		/// <description> 项目符号格式（如果适用） </description>
		/// </item>
		/// </list>
		/// </remarks>
		void ApplyTextFormatting(NETOP.Shape shp);

		/// <summary>
		/// 应用文本格式化到指定形状（抽象接口版本）
		/// </summary>
		/// <param name="shape"> 要格式化的抽象形状对象，不能为 null </param>
		/// <remarks> 此方法内部会将抽象接口转换为 NetOffice 对象，然后调用 NetOffice 版本的方法。 </remarks>
		void ApplyTextFormatting(IShape shape);
	}
}

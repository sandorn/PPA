using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 表格格式化辅助接口 提供表格的高性能格式化功能，包括样式、边框、字体等设置
	/// </summary>
	/// <remarks>
	/// 此接口定义了表格格式化的接口，通过依赖注入使用，便于测试和扩展。 格式化参数从配置文件（ <see cref="IFormattingConfig" />）中读取。 实现类使用高性能的批量操作方式，避免逐单元格设置导致的性能问题。
	/// <para> <strong> 接口版本说明： </strong> </para>
	/// <list type="bullet">
	/// <item>
	/// <description>
	/// <strong> NetOffice 版本 </strong>（ <see cref="FormatTables(NETOP.Table)" />）：
	/// 提供完整的格式化功能，包括样式、边框、字体、数字格式等所有功能。 这是主要使用的版本，功能最完整。
	/// </description>
	/// </item>
	/// <item>
	/// <description>
	/// <strong> 抽象接口版本 </strong>（ <see cref="FormatTables(ITable)" />）： 内部通过适配器模式转换为 NetOffice
	/// 对象后调用 NetOffice 版本。 主要用于与抽象接口系统的集成，功能覆盖度依赖于抽象接口的实现。
	/// </description>
	/// </item>
	/// </list>
	/// </remarks>
	public interface ITableFormatHelper
	{
		/// <summary>
		/// 对表格进行格式化（NetOffice 版本）
		/// </summary>
		/// <param name="tbl"> 要格式化的 NetOffice 表格对象，不能为 null </param>
		/// <remarks>
		/// 此方法会应用以下格式化设置：
		/// <list type="bullet">
		/// <item>
		/// <description> 表格样式（表头、数据行的背景色和字体） </description>
		/// </item>
		/// <item>
		/// <description> 边框样式（表头和数据行的边框宽度和颜色） </description>
		/// </item>
		/// <item>
		/// <description> 字体属性（名称、大小、颜色等） </description>
		/// </item>
		/// <item>
		/// <description> 数字格式（自动编号、小数位数、负数颜色等） </description>
		/// </item>
		/// </list>
		/// </remarks>
		void FormatTables(NETOP.Table tbl);

		/// <summary>
		/// 对表格进行格式化（抽象接口版本）
		/// </summary>
		/// <param name="tbl"> 要格式化的抽象表格对象，不能为 null </param>
		/// <remarks> 此方法内部会将抽象接口转换为 NetOffice 对象，然后调用 NetOffice 版本的方法。 </remarks>
		void FormatTables(ITable tbl);
	}
}

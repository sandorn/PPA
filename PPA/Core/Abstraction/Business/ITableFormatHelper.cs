using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 表格格式化辅助接口 提供表格的高性能格式化功能，包括样式、边框、字体等设置
	/// </summary>
	/// <remarks>
	/// 此接口定义了表格格式化的接口，通过依赖注入使用，便于测试和扩展。 格式化参数从配置文件（ <see cref="IFormattingConfig" />）中读取。 实现类使用高性能的批量操作方式，避免逐单元格设置导致的性能问题。
	/// </remarks>
	public interface ITableFormatHelper
	{
		/// <summary>
		/// 对表格进行格式化
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
	}
}

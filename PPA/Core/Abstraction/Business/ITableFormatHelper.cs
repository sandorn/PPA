using NETOP = NetOffice.PowerPointApi;
using PPA.Core.Abstraction.Presentation;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 表格格式化辅助接口
	/// 注意：当前使用 NetOffice 类型，后续阶段将改为使用平台抽象接口
	/// </summary>
	public interface ITableFormatHelper
	{
		/// <summary>
		/// 对表格进行格式化
		/// </summary>
		/// <param name="tbl">要格式化的表格对象</param>
		void FormatTables(NETOP.Table tbl);

		/// <summary>
		/// 对表格进行格式化（抽象接口版本）
		/// </summary>
		/// <param name="tbl">要格式化的抽象表格对象</param>
		void FormatTables(ITable tbl);
	}
}


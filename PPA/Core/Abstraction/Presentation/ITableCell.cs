namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示表格单元格的抽象接口
	/// </summary>
	/// <remarks>
	/// 此接口封装了 PowerPoint 中的表格单元格对象，提供了统一的单元格访问接口。 单元格是表格的基本单位，可以包含文本、数字等内容。 实现类： <see cref="PPA.Core.Adapters.PowerPoint.PowerPointTableCell" />。
	/// </remarks>
	public interface ITableCell:IComWrapper
	{
		/// <summary>
		/// 获取单元格所在的行索引（1-based）
		/// </summary>
		/// <value> 行索引，从 1 开始 </value>
		/// <remarks> 行索引从 1 开始，第一行的索引为 1。 </remarks>
		int RowIndex { get; }

		/// <summary>
		/// 获取单元格所在的列索引（1-based）
		/// </summary>
		/// <value> 列索引，从 1 开始 </value>
		/// <remarks> 列索引从 1 开始，第一列的索引为 1。 </remarks>
		int ColumnIndex { get; }

		/// <summary>
		/// 获取单元格中的文本范围
		/// </summary>
		/// <returns> 文本范围对象，如果单元格中没有文本则返回 null </returns>
		/// <remarks> 此方法返回单元格中的文本内容，可以用于读取或修改单元格文本。 如果单元格为空或没有文本内容，则返回 null。 </remarks>
		ITextRange GetTextRange();
	}
}

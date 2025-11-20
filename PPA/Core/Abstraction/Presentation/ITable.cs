namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示表格对象的抽象接口
	/// </summary>
	/// <remarks> 此接口封装了 PowerPoint 中的表格对象，提供了统一的表格访问接口。 表格由多个单元格组成，可以包含文本、数字等内容。 实现类： <see cref="PPA.Core.Adapters.PowerPoint.PowerPointTable" />。 </remarks>
	public interface ITable:IComWrapper
	{
		/// <summary>
		/// 获取表格的行数
		/// </summary>
		/// <value> 表格的行数，至少为 1 </value>
		int RowCount { get; }

		/// <summary>
		/// 获取表格的列数
		/// </summary>
		/// <value> 表格的列数，至少为 1 </value>
		int ColumnCount { get; }

		/// <summary>
		/// 根据行列索引获取单元格（1-based）
		/// </summary>
		/// <param name="row"> 行号，范围为 [1, <see cref="RowCount" />] </param>
		/// <param name="column"> 列号，范围为 [1, <see cref="ColumnCount" />] </param>
		/// <returns> 单元格对象，如果索引超出范围则可能返回 null 或抛出异常 </returns>
		/// <exception cref="ArgumentOutOfRangeException"> 当 row 或 column 超出有效范围时抛出 </exception>
		/// <remarks> 行号和列号都从 1 开始，第一行第一列的单元格为 (1, 1)。 </remarks>
		ITableCell GetCell(int row,int column);
	}
}

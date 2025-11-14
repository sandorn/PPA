namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示表格对象
	/// </summary>
	public interface ITable : IComWrapper
	{
		/// <summary>
		/// 行数
		/// </summary>
		int RowCount { get; }

		/// <summary>
		/// 列数
		/// </summary>
		int ColumnCount { get; }

		/// <summary>
		/// 根据行列索引获取单元格（1-based），如不存在返回 null
		/// </summary>
		/// <param name="row">行号</param>
		/// <param name="column">列号</param>
		/// <returns>单元格对象</returns>
		ITableCell GetCell(int row, int column);
	}
}



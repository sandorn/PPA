namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表格单元格接口
	/// </summary>
	public interface ITableCell : IComWrapper
	{
		/// <summary>
		/// 对应的行索引（1-based）
		/// </summary>
		int RowIndex { get; }

		/// <summary>
		/// 对应的列索引（1-based）
		/// </summary>
		int ColumnIndex { get; }

		/// <summary>
		/// 单元格中的文本范围，没有文本时返回 null
		/// </summary>
		ITextRange GetTextRange();
	}
}



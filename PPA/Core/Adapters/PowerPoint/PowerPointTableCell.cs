using System;
using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 表格单元格适配器
	/// </summary>
	public sealed class PowerPointTableCell : ITableCell, IComWrapper<NETOP.Cell>
	{
		public ITable Table { get; }
		public NETOP.Cell NativeObject { get; }
		object IComWrapper.NativeObject => NativeObject;

		public int RowIndex { get; }
		public int ColumnIndex { get; }

		public PowerPointTableCell(ITable table, NETOP.Cell cell, int row, int column)
		{
			Table = table ?? throw new ArgumentNullException(nameof(table));
			NativeObject = cell ?? throw new ArgumentNullException(nameof(cell));
			RowIndex = row;
			ColumnIndex = column;
		}

		public ITextRange GetTextRange()
		{
			try
			{
				var shape = NativeObject?.Shape;
				var tr = shape?.TextFrame?.TextRange;
				return tr != null ? new PowerPointTextRange(Table is PowerPointTable pptTable ? pptTable.ParentShape : null, tr) : null;
			} catch { return null; }
		}
	}
}



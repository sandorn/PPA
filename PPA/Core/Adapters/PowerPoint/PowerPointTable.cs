using System;
using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 表格适配器
	/// </summary>
	public sealed class PowerPointTable : ITable, IComWrapper<NETOP.Table>
	{
		public IShape ParentShape { get; }
		public NETOP.Table NativeObject { get; }
		object IComWrapper.NativeObject => NativeObject;

		public int RowCount => SafeGet(() => NativeObject?.Rows?.Count ?? 0, 0);
		public int ColumnCount => SafeGet(() => NativeObject?.Columns?.Count ?? 0, 0);

		public PowerPointTable(IShape parent, NETOP.Table table)
		{
			ParentShape = parent ?? throw new ArgumentNullException(nameof(parent));
			NativeObject = table ?? throw new ArgumentNullException(nameof(table));
		}

		public ITableCell GetCell(int row, int column)
		{
			try
			{
				var cell = NativeObject.Cell(row,column);
				return cell != null ? new PowerPointTableCell(this,cell,row,column) : null;
			} catch { return null; }
		}

		private static T SafeGet<T>(Func<T> getter, T fallback)
		{
			try { return getter(); } catch { return fallback; }
		}
	}
}



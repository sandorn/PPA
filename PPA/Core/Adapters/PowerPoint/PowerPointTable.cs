using PPA.Core.Abstraction.Presentation;
using System;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 表格适配器
	/// </summary>
	public sealed class PowerPointTable(IShape parent,NETOP.Table table):ITable, IComWrapper<NETOP.Table>
	{
		public IShape ParentShape { get; } = parent??throw new ArgumentNullException(nameof(parent));
		public NETOP.Table NativeObject { get; } = table??throw new ArgumentNullException(nameof(table));
		object IComWrapper.NativeObject => NativeObject;

		public int RowCount => ExHandler.SafeGet(() => NativeObject?.Rows?.Count??0,0);
		public int ColumnCount => ExHandler.SafeGet(() => NativeObject?.Columns?.Count??0,0);

		public ITableCell GetCell(int row,int column)
		{
			try
			{
				var cell = NativeObject.Cell(row,column);
				return cell!=null ? new PowerPointTableCell(this,cell,row,column) : null;
			} catch { return null; }
		}
	}
}

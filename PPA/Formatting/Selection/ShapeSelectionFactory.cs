using PPA.Core;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using PPA.Core.Logging;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting.Selection
{
	/// <summary>
	/// 表示可以被迭代处理的 PowerPoint 形状选区
	/// </summary>
	internal interface IShapeSelection:IEnumerable<NETOP.Shape>
	{
		int Count { get; }
	}

	internal sealed class EmptyShapeSelection:IShapeSelection
	{
		public static readonly EmptyShapeSelection Instance = new();

		public int Count => 0;

		public IEnumerator<NETOP.Shape> GetEnumerator()
		{
			yield break;
		}

		IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
	}

	internal sealed class SingleShapeSelection(NETOP.Shape shape):IShapeSelection
	{
		private readonly NETOP.Shape _shape = shape;

		public int Count => _shape!=null ? 1 : 0;

		public IEnumerator<NETOP.Shape> GetEnumerator()
		{
			if(_shape!=null)
			{
				yield return _shape;
			}
		}

		IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
	}

	internal sealed class ShapeRangeSelection:IShapeSelection
	{
		private static readonly ILogger Logger = LoggerProvider.GetLogger();
		private readonly NETOP.ShapeRange _shapeRange;

		public ShapeRangeSelection(NETOP.ShapeRange shapeRange) => _shapeRange=shapeRange;

		public int Count => ExHandler.SafeGet(() => _shapeRange.Count,0);

		public IEnumerator<NETOP.Shape> GetEnumerator()
		{
			if(_shapeRange==null)
			{
				yield break;
			}

			var total = Count;
			if(total<=0)
			{
				yield break;
			}

			for(int i = 1;i<=total;i++)
			{
				NETOP.Shape shape = null;
				try
				{
					shape=_shapeRange[i];
				} catch(System.Exception ex)
				{
					Logger.LogDebug($"ShapeRangeSelection: 获取索引 {i} 的形状失败: {ex.Message}");
				}

				if(shape!=null)
				{
					yield return shape;
				}
			}
		}

		IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
	}

	internal sealed class AbstractShapeSelection:IShapeSelection
	{
		private readonly IShape _shape;

		public AbstractShapeSelection(IShape shape) => _shape=shape;

		public int Count => _shape!=null ? 1 : 0;

		public IEnumerator<NETOP.Shape> GetEnumerator()
		{
			if(_shape==null)
			{
				yield break;
			}

			var pptShape = AdapterUtils.UnwrapShape(_shape);
			if(pptShape!=null)
			{
				yield return pptShape;
			}
		}

		IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
	}

	internal sealed class AbstractShapeCollectionSelection:IShapeSelection
	{
		private readonly IEnumerable<IShape> _shapes;

		public AbstractShapeCollectionSelection(IEnumerable<IShape> shapes) => _shapes=shapes??Enumerable.Empty<IShape>();

		public int Count => _shapes.Count();

		public IEnumerator<NETOP.Shape> GetEnumerator()
		{
			foreach(var abstractShape in _shapes)
			{
				var pptShape = AdapterUtils.UnwrapShape(abstractShape);
				if(pptShape!=null)
				{
					yield return pptShape;
				}
			}
		}

		IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
	}

	internal static class ShapeSelectionFactory
	{
		public static IShapeSelection Create(object selection)
		{
			if(selection==null)
			{
				return null;
			}

			return selection switch
			{
				NETOP.Shape shape => new SingleShapeSelection(shape),
				NETOP.ShapeRange shapeRange => new ShapeRangeSelection(shapeRange),
				IShape abstractShape => new AbstractShapeSelection(abstractShape),
				IEnumerable<IShape> abstractShapes => new AbstractShapeCollectionSelection(abstractShapes),
				_ => null
			};
		}
	}
}

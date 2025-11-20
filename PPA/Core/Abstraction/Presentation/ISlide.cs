using System.Collections.Generic;

namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示幻灯片的抽象接口
	/// </summary>
	/// <remarks>
	/// 此接口封装了 PowerPoint 中的幻灯片对象，提供了统一的幻灯片访问接口。 幻灯片是演示文稿的基本单元，包含多个形状（文本框、表格、图表、图片等）。 实现类： <see cref="PPA.Core.Adapters.PowerPoint.PowerPointSlide" />。
	/// </remarks>
	public interface ISlide:IComWrapper
	{
		/// <summary>
		/// 获取幻灯片所属的应用程序
		/// </summary>
		/// <value> 应用程序对象，不会为 null </value>
		IApplication Application { get; }

		/// <summary>
		/// 获取幻灯片所属的演示文稿
		/// </summary>
		/// <value> 演示文稿对象，不会为 null </value>
		IPresentation Presentation { get; }

		/// <summary>
		/// 获取幻灯片标题
		/// </summary>
		/// <value> 幻灯片标题，如果幻灯片没有标题则返回空字符串 </value>
		/// <remarks> 幻灯片标题通常来自标题占位符中的文本内容。 </remarks>
		string Title { get; }

		/// <summary>
		/// 获取幻灯片序号（1-based）
		/// </summary>
		/// <value> 幻灯片在演示文稿中的序号，从 1 开始 </value>
		/// <remarks> 幻灯片序号从 1 开始，第一张幻灯片的序号为 1。 </remarks>
		int SlideIndex { get; }

		/// <summary>
		/// 获取幻灯片中的所有形状集合
		/// </summary>
		/// <value> 形状集合，按 Z-Order 顺序排列，如果没有形状则返回空集合（不会返回 null） </value>
		/// <remarks> 此属性返回幻灯片中的所有形状，包括文本框、表格、图表、图片、自选图形等。 形状按 Z-Order（层叠顺序）排列，后添加的形状通常在上层。 </remarks>
		IReadOnlyList<IShape> Shapes { get; }

		/// <summary>
		/// 根据名称查找形状
		/// </summary>
		/// <param name="name"> 要查找的形状名称，不能为 null 或空字符串 </param>
		/// <returns> 找到的形状对象，如果不存在则返回 null </returns>
		/// <remarks> 此方法在幻灯片的所有形状中查找指定名称的形状。 形状名称区分大小写，如果存在多个同名形状，则返回第一个找到的形状。 </remarks>
		IShape FindShapeByName(string name);
	}
}

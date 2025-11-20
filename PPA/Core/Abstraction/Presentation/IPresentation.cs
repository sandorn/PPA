using System.Collections.Generic;

namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示演示文稿文档的抽象接口
	/// </summary>
	/// <remarks>
	/// 此接口封装了 PowerPoint 中的演示文稿对象，提供了统一的演示文稿访问接口。 演示文稿是 PowerPoint 的文档单位，包含多个幻灯片。 实现类： <see cref="PPA.Core.Adapters.PowerPoint.PowerPointPresentation" />。
	/// </remarks>
	public interface IPresentation:IComWrapper
	{
		/// <summary>
		/// 获取演示文稿所属的应用程序实例
		/// </summary>
		/// <value> 应用程序对象，不会为 null </value>
		IApplication Application { get; }

		/// <summary>
		/// 获取演示文稿名称
		/// </summary>
		/// <value> 演示文稿的文件名（不含路径），如果演示文稿未保存则可能返回临时名称 </value>
		/// <remarks> 演示文稿名称通常是文件名（不含扩展名），例如 "MyPresentation"。 如果演示文稿是新创建的且未保存，则可能返回临时名称。 </remarks>
		string Name { get; }

		/// <summary>
		/// 获取演示文稿中的幻灯片数量
		/// </summary>
		/// <value> 幻灯片数量，至少为 1（演示文稿至少包含一张幻灯片） </value>
		int SlideCount { get; }

		/// <summary>
		/// 获取演示文稿中的所有幻灯片集合
		/// </summary>
		/// <value> 幻灯片集合，按显示顺序排列，不会为 null </value>
		/// <remarks> 此属性返回演示文稿中的所有幻灯片，按在演示文稿中的顺序排列。 </remarks>
		IReadOnlyList<ISlide> Slides { get; }

		/// <summary>
		/// 按索引获取幻灯片（1-based）
		/// </summary>
		/// <param name="index"> 幻灯片索引，范围为 [1, <see cref="SlideCount" />] </param>
		/// <returns> 幻灯片对象，如果索引超出范围则可能返回 null 或抛出异常 </returns>
		/// <exception cref="ArgumentOutOfRangeException">
		/// 当 index 小于 1 或大于 <see cref="SlideCount" /> 时抛出
		/// </exception>
		/// <remarks> 幻灯片索引从 1 开始，第一张幻灯片的索引为 1。 </remarks>
		ISlide GetSlide(int index);
	}
}

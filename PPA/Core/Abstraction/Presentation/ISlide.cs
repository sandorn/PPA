using System.Collections.Generic;

namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示幻灯片的抽象接口
	/// </summary>
	public interface ISlide : IComWrapper
	{
		/// <summary>
		/// 所属应用程序
		/// </summary>
		IApplication Application { get; }

		/// <summary>
		/// 所属演示文稿
		/// </summary>
		IPresentation Presentation { get; }

		/// <summary>
		/// 幻灯片标题（如果存在）
		/// </summary>
		string Title { get; }

		/// <summary>
		/// 幻灯片序号（1-based）
		/// </summary>
		int SlideIndex { get; }

		/// <summary>
		/// 幻灯片中的所有形状集合
		/// </summary>
		IReadOnlyList<IShape> Shapes { get; }

		/// <summary>
		/// 根据名称查找形状，不存在时返回 null
		/// </summary>
		/// <param name="name">形状名称</param>
		/// <returns>IShape 或 null</returns>
		IShape FindShapeByName(string name);
	}
}



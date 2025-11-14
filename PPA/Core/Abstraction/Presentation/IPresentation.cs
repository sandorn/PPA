using System.Collections.Generic;

namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示演示文稿文档的抽象接口
	/// </summary>
	public interface IPresentation : IComWrapper
	{
		/// <summary>
		/// 所属应用程序实例
		/// </summary>
		IApplication Application { get; }

		/// <summary>
		/// 演示文稿名称
		/// </summary>
		string Name { get; }

		/// <summary>
		/// 幻灯片数量
		/// </summary>
		int SlideCount { get; }

		/// <summary>
		/// 获取演示文稿中的所有幻灯片，按显示顺序排列
		/// </summary>
		IReadOnlyList<ISlide> Slides { get; }

		/// <summary>
		/// 按索引获取幻灯片（1-based）
		/// </summary>
		/// <param name="index">幻灯片索引，范围为 [1, SlideCount]</param>
		/// <returns>幻灯片对象</returns>
		ISlide GetSlide(int index);
	}
}



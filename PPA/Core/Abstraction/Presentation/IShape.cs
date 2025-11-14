namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示幻灯片中的形状
	/// </summary>
	public interface IShape : IComWrapper
	{
		/// <summary>
		/// 所属应用程序
		/// </summary>
		IApplication Application { get; }

		/// <summary>
		/// 所属幻灯片
		/// </summary>
		ISlide Slide { get; }

		/// <summary>
		/// 形状名称
		/// </summary>
		string Name { get; }

		/// <summary>
		/// 形状类型标识（平台相关），用于快速判断类型
		/// </summary>
		int ShapeType { get; }

		/// <summary>
		/// 是否包含文本框
		/// </summary>
		bool HasText { get; }

		/// <summary>
		/// 是否包含表格
		/// </summary>
		bool HasTable { get; }

		/// <summary>
		/// 是否包含图表
		/// </summary>
		bool HasChart { get; }

		/// <summary>
		/// 获取文本范围（如不存在返回 null）
		/// </summary>
		ITextRange GetTextRange();

		/// <summary>
		/// 获取表格对象（如不存在返回 null）
		/// </summary>
		ITable GetTable();

		/// <summary>
		/// 获取图表对象（如不存在返回 null）
		/// </summary>
		IChart GetChart();
	}
}



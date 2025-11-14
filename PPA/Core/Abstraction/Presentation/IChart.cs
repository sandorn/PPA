namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 图表抽象接口
	/// </summary>
	public interface IChart : IComWrapper
	{
		/// <summary>
		/// 图表所属的形状
		/// </summary>
		IShape ParentShape { get; }

		/// <summary>
		/// 图表类型标识（平台相关）
		/// </summary>
		int ChartType { get; }

		/// <summary>
		/// 将图表应用预定义样式
		/// </summary>
		/// <param name="styleId">样式标识</param>
		void ApplyPredefinedStyle(string styleId);
	}
}



namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示图表对象的抽象接口
	/// </summary>
	/// <remarks> 此接口封装了 PowerPoint 中的图表对象，提供了统一的图表访问接口。 图表可以是柱状图、折线图、饼图等多种类型。 实现类： <see cref="PPA.Core.Adapters.PowerPoint.PowerPointChart" />。 </remarks>
	public interface IChart:IComWrapper
	{
		/// <summary>
		/// 获取图表所属的形状
		/// </summary>
		/// <value> 包含此图表的形状对象，不会为 null </value>
		/// <remarks> 图表总是包含在一个形状中，通过此属性可以访问包含图表的形状对象。 </remarks>
		IShape ParentShape { get; }

		/// <summary>
		/// 获取图表类型标识（平台相关）
		/// </summary>
		/// <value> 图表类型标识，用于标识图表的类型（如柱状图、折线图、饼图等） </value>
		/// <remarks> 此属性返回平台相关的图表类型标识，不同平台的值可能不同。 用于快速判断图表类型，避免频繁调用底层 COM 对象。 </remarks>
		int ChartType { get; }

		/// <summary>
		/// 将图表应用预定义样式
		/// </summary>
		/// <param name="styleId"> 样式标识，例如 "Style1"、"Style2" 等，不能为 null 或空字符串 </param>
		/// <remarks> 此方法会将图表应用 PowerPoint 中预定义的图表样式。 样式标识通常是字符串格式，如 "Style1"、"Style2" 等。 </remarks>
		void ApplyPredefinedStyle(string styleId);
	}
}

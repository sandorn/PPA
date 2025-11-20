namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示幻灯片中的形状的抽象接口
	/// </summary>
	/// <remarks> 此接口封装了 PowerPoint 中的形状对象，提供了统一的形状访问接口。 形状可以是文本框、表格、图表、图片、自选图形等。 实现类： <see cref="PPA.Core.Adapters.PowerPoint.PowerPointShape" />。 </remarks>
	public interface IShape:IComWrapper
	{
		/// <summary>
		/// 获取形状所属的应用程序
		/// </summary>
		/// <value> 应用程序对象，不会为 null </value>
		IApplication Application { get; }

		/// <summary>
		/// 获取形状所属的幻灯片
		/// </summary>
		/// <value> 幻灯片对象，不会为 null </value>
		ISlide Slide { get; }

		/// <summary>
		/// 获取形状名称
		/// </summary>
		/// <value> 形状的名称，PowerPoint 中每个形状都有一个唯一名称 </value>
		/// <remarks> 形状名称通常由 PowerPoint 自动生成（如 "Rectangle 1"、"Text Box 2"）， 用户也可以手动修改形状名称。 </remarks>
		string Name { get; }

		/// <summary>
		/// 获取形状类型标识（平台相关）
		/// </summary>
		/// <value> 形状类型标识，用于快速判断形状类型（如矩形、文本框、表格等） </value>
		/// <remarks>
		/// 此属性返回平台相关的形状类型标识，不同平台的值可能不同。 用于快速判断形状类型，避免频繁调用 <see cref="HasText" />、
		/// <see cref="HasTable" /> 等属性。
		/// </remarks>
		int ShapeType { get; }

		/// <summary>
		/// 获取是否包含文本框
		/// </summary>
		/// <value> 如果形状包含文本框则为 true，否则为 false </value>
		/// <remarks> 大多数形状都可以包含文本，此属性用于快速判断形状是否包含文本内容。 </remarks>
		bool HasText { get; }

		/// <summary>
		/// 获取是否包含表格
		/// </summary>
		/// <value> 如果形状包含表格则为 true，否则为 false </value>
		/// <remarks> 只有表格形状（Table）才包含表格对象。 </remarks>
		bool HasTable { get; }

		/// <summary>
		/// 获取是否包含图表
		/// </summary>
		/// <value> 如果形状包含图表则为 true，否则为 false </value>
		/// <remarks> 只有图表形状（Chart）才包含图表对象。 </remarks>
		bool HasChart { get; }

		/// <summary>
		/// 获取文本范围对象
		/// </summary>
		/// <returns> 文本范围对象，如果形状不包含文本则返回 null </returns>
		/// <remarks>
		/// 此方法返回形状中的文本内容，可以用于读取或修改文本。 如果形状不包含文本（ <see cref="HasText" /> 为 false），则返回 null。
		/// </remarks>
		ITextRange GetTextRange();

		/// <summary>
		/// 获取表格对象
		/// </summary>
		/// <returns> 表格对象，如果形状不包含表格则返回 null </returns>
		/// <remarks>
		/// 此方法返回形状中的表格对象，可以用于访问和修改表格内容。 如果形状不包含表格（ <see cref="HasTable" /> 为 false），则返回 null。
		/// </remarks>
		ITable GetTable();

		/// <summary>
		/// 获取图表对象
		/// </summary>
		/// <returns> 图表对象，如果形状不包含图表则返回 null </returns>
		/// <remarks>
		/// 此方法返回形状中的图表对象，可以用于访问和修改图表数据、样式等。 如果形状不包含图表（ <see cref="HasChart" /> 为 false），则返回 null。
		/// </remarks>
		IChart GetChart();
	}
}

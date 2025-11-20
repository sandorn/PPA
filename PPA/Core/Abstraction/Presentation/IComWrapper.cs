namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 统一的 COM 对象包装接口 提供对底层原生 COM 对象的访问，实现抽象接口与具体 COM 对象的解耦
	/// </summary>
	/// <remarks>
	/// 此接口是所有抽象接口的基接口，提供了访问底层原生 COM 对象的能力。 通过此接口，可以在需要时访问底层的 NetOffice 或原生 COM 对象。 所有抽象接口（如
	/// <see cref="IApplication" />、 <see cref="IShape" /> 等）都继承自此接口。
	/// </remarks>
	public interface IComWrapper
	{
		/// <summary>
		/// 获取底层原生 COM 对象
		/// </summary>
		/// <value> 底层原生 COM 对象，通常是 NetOffice 包装的对象或原生 COM 对象 </value>
		/// <remarks> 此属性返回底层的原生 COM 对象，可以用于需要直接访问 COM 接口的场景。 大多数情况下，应优先使用抽象接口的方法，而不是直接访问原生对象。 </remarks>
		object NativeObject { get; }
	}

	/// <summary>
	/// 泛型版本的 COM 包装接口 为调用方提供强类型的原生对象访问，避免类型转换
	/// </summary>
	/// <typeparam name="TNative">
	/// 原生对象类型，例如 <c> NetOffice.PowerPointApi.Application </c>、 <c> NetOffice.PowerPointApi.Shape
	/// </c> 等
	/// </typeparam>
	/// <remarks> 此接口继承自 <see cref="IComWrapper" />，提供了强类型的原生对象访问。 使用此接口可以避免类型转换，提高代码的类型安全性。 </remarks>
	public interface IComWrapper<out TNative>:IComWrapper
	{
		/// <summary>
		/// 获取强类型的原生对象
		/// </summary>
		/// <value> 强类型的原生 COM 对象，类型为 <typeparamref name="TNative" /> </value>
		/// <remarks>
		/// 此属性返回强类型的原生对象，避免了从 <see cref="IComWrapper.NativeObject" /> 进行类型转换。 如果底层对象不是
		/// <typeparamref name="TNative" /> 类型，则可能返回 null 或抛出异常。
		/// </remarks>
		new TNative NativeObject { get; }
	}
}

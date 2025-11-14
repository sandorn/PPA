namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 统一的 COM 对象包装接口，提供对底层原生对象的访问
	/// </summary>
	public interface IComWrapper
	{
		/// <summary>
		/// 获取底层原生 COM 对象
		/// </summary>
		object NativeObject { get; }
	}

	/// <summary>
	/// 泛型版本的 COM 包装接口，为调用方提供强类型的原生对象访问
	/// </summary>
	/// <typeparam name="TNative">原生对象类型</typeparam>
	public interface IComWrapper<out TNative> : IComWrapper
	{
		/// <summary>
		/// 获取强类型的原生对象
		/// </summary>
		new TNative NativeObject { get; }
	}
}



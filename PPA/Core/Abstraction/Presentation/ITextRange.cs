namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 文本范围抽象接口
	/// </summary>
	public interface ITextRange : IComWrapper
	{
		/// <summary>
		/// 文本内容
		/// </summary>
		string Text { get; set; }

		/// <summary>
		/// 应用预定义格式
		/// </summary>
		/// <param name="formatId">格式标识，例如 "TitleText"、"ContentText"</param>
		void ApplyPredefinedFormat(string formatId);
	}
}



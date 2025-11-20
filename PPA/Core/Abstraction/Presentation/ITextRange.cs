namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示文本范围的抽象接口
	/// </summary>
	/// <remarks>
	/// 此接口封装了 PowerPoint 中的文本范围对象，提供了统一的文本访问接口。 文本范围可以是一个字符、一个单词、一个段落或整个文本框的内容。 实现类： <see cref="PPA.Core.Adapters.PowerPoint.PowerPointTextRange" />。
	/// </remarks>
	public interface ITextRange:IComWrapper
	{
		/// <summary>
		/// 获取或设置文本内容
		/// </summary>
		/// <value> 文本内容，如果文本范围为空则返回空字符串 </value>
		/// <remarks> 此属性可以用于读取或修改文本范围中的文本内容。 设置文本内容会替换当前文本范围中的所有文本。 </remarks>
		string Text { get; set; }

		/// <summary>
		/// 应用预定义格式到文本范围
		/// </summary>
		/// <param name="formatId"> 格式标识，例如 "TitleText"、"ContentText" 等，不能为 null 或空字符串 </param>
		/// <remarks>
		/// 此方法会将文本范围应用 PowerPoint 中预定义的文本格式。 格式标识通常是字符串格式，如 "TitleText"、"ContentText" 等。 如果指定的格式标识不存在，此方法可能不会产生任何效果或抛出异常。
		/// </remarks>
		void ApplyPredefinedFormat(string formatId);
	}
}

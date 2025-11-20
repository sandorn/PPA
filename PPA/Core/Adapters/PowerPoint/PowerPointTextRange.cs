using PPA.Core.Abstraction.Presentation;
using System;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 文本范围适配器
	/// </summary>
	public sealed class PowerPointTextRange(IShape parentShape,NETOP.TextRange textRange):ITextRange, IComWrapper<NETOP.TextRange>
	{
		public IShape ParentShape { get; } = parentShape;
		public NETOP.TextRange NativeObject { get; } = textRange??throw new ArgumentNullException(nameof(textRange));
		object IComWrapper.NativeObject => NativeObject;

		public string Text
		{
			get => ExHandler.SafeGet(() => NativeObject?.Text,string.Empty);
			set => ExHandler.SafeSet(() => NativeObject!.Text=value??string.Empty);
		}

		public void ApplyPredefinedFormat(string formatId)
		{
			// 预留：根据 formatId 应用字体、字号、颜色等
		}
	}
}

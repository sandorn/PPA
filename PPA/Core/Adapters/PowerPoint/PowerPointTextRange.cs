using System;
using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 文本范围适配器
	/// </summary>
	public sealed class PowerPointTextRange : ITextRange, IComWrapper<NETOP.TextRange>
	{
		public IShape ParentShape { get; }
		public NETOP.TextRange NativeObject { get; }
		object IComWrapper.NativeObject => NativeObject;

		public PowerPointTextRange(IShape parentShape, NETOP.TextRange textRange)
		{
			ParentShape = parentShape;
			NativeObject = textRange ?? throw new ArgumentNullException(nameof(textRange));
		}

		public string Text
		{
			get => SafeGet(() => NativeObject?.Text, string.Empty);
			set
			{
				try { if(NativeObject!=null) NativeObject.Text = value ?? string.Empty; } catch { /* ignore */ }
			}
		}

		public void ApplyPredefinedFormat(string formatId)
		{
			// 预留：根据 formatId 应用字体、字号、颜色等
		}

		private static T SafeGet<T>(Func<T> getter, T fallback)
		{
			try { return getter(); } catch { return fallback; }
		}
	}
}



using System;
using System.Text;

namespace Project.Utilities
{
	public static class ExFormatter
	{
		#region Public Methods

		public static string FormatFullException(Exception ex)
		{
			if(ex == null) return string.Empty;

			var sb = new StringBuilder();
			AppendExceptionDetails(sb,ex,depth: 0);
			return sb.ToString();
		}

		#endregion Public Methods

		#region Private Methods

		private static void AppendExceptionDetails(StringBuilder sb,Exception ex,int depth)
		{
			if(depth > 0) sb.Append('\n').Append(' ',depth * 2);

			sb.Append($"[{ex.GetType().Name}] {ex.Message}");
			sb.Append($"\n{"HResult:",-10} 0x{ex.HResult:X8}");

			if(!string.IsNullOrWhiteSpace(ex.StackTrace))
			{
				sb.Append($"\n{"Stack Trace:",-10}");
				sb.Append(FormatStackTrace(ex.StackTrace));
			}

			if(ex.InnerException != null)
			{
				sb.Append($"\n{"Inner:",-10}");
				AppendExceptionDetails(sb,ex.InnerException,depth + 1);
			}
		}

		private static string FormatStackTrace(string stackTrace)
		{
			var lines = stackTrace.Split(new[] { '\r','\n' },StringSplitOptions.RemoveEmptyEntries);
			return "\n          " + string.Join("\n          ",lines);
		}

		#endregion Private Methods
	}
}
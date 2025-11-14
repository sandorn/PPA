using System;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// Office 常用命令常量
	/// 提供常用的 MSO 命令名称和命令 ID
	/// </summary>
	public static class OfficeCommands
	{
		#region 剪贴板命令

		/// <summary>
		/// 复制
		/// </summary>
		public const string Copy = "Copy";

		/// <summary>
		/// 剪切
		/// </summary>
		public const string Cut = "Cut";

		/// <summary>
		/// 粘贴
		/// </summary>
		public const string Paste = "Paste";

		/// <summary>
		/// 选择性粘贴
		/// </summary>
		public const string PasteSpecial = "PasteSpecial";

		#endregion

		#region 格式命令

		/// <summary>
		/// 加粗
		/// </summary>
		public const string Bold = "Bold";

		/// <summary>
		/// 斜体
		/// </summary>
		public const string Italic = "Italic";

		/// <summary>
		/// 下划线
		/// </summary>
		public const string Underline = "Underline";

		/// <summary>
		/// 增大字号
		/// </summary>
		public const string FontSizeIncrease = "FontSizeIncrease";

		/// <summary>
		/// 减小字号
		/// </summary>
		public const string FontSizeDecrease = "FontSizeDecrease";

		#endregion

		#region 幻灯片命令

		/// <summary>
		/// 新建幻灯片
		/// </summary>
		public const string SlideNew = "SlideNew";

		/// <summary>
		/// 删除幻灯片
		/// </summary>
		public const string SlideDelete = "SlideDelete";

		/// <summary>
		/// 复制幻灯片
		/// </summary>
		public const string SlideDuplicate = "SlideDuplicate";

		/// <summary>
		/// 从开始放映
		/// </summary>
		public const string SlideShowFromBeginning = "SlideShowFromBeginning";

		/// <summary>
		/// 从当前幻灯片放映
		/// </summary>
		public const string SlideShowFromCurrent = "SlideShowFromCurrent";

		#endregion

		#region 文件命令

		/// <summary>
		/// 保存
		/// </summary>
		public const string FileSave = "FileSave";

		/// <summary>
		/// 另存为
		/// </summary>
		public const string FileSaveAs = "FileSaveAs";

		/// <summary>
		/// 打开
		/// </summary>
		public const string FileOpen = "FileOpen";

		/// <summary>
		/// 新建
		/// </summary>
		public const string FileNew = "FileNew";

		#endregion

		#region 编辑命令

		/// <summary>
		/// 撤销
		/// </summary>
		public const string Undo = "Undo";

		/// <summary>
		/// 重做
		/// </summary>
		public const string Redo = "Redo";

		/// <summary>
		/// 全选
		/// </summary>
		public const string SelectAll = "SelectAll";

		/// <summary>
		/// 查找
		/// </summary>
		public const string Find = "Find";

		/// <summary>
		/// 替换
		/// </summary>
		public const string Replace = "Replace";

		#endregion

		#region 形状对齐命令

		/// <summary>
		/// 形状左对齐（Smart 命令）
		/// </summary>
		public const string ObjectsAlignLeftSmart = "ObjectsAlignLeftSmart";

		/// <summary>
		/// 形状水平居中（Smart 命令）
		/// </summary>
		public const string ObjectsAlignCenterHorizontalSmart = "ObjectsAlignCenterHorizontalSmart";

		/// <summary>
		/// 形状右对齐（Smart 命令）
		/// </summary>
		public const string ObjectsAlignRightSmart = "ObjectsAlignRightSmart";

		/// <summary>
		/// 形状顶对齐（Smart 命令）
		/// </summary>
		public const string ObjectsAlignTopSmart = "ObjectsAlignTopSmart";

		/// <summary>
		/// 形状垂直居中（Smart 命令）
		/// </summary>
		public const string ObjectsAlignMiddleVerticalSmart = "ObjectsAlignMiddleVerticalSmart";

		/// <summary>
		/// 形状底对齐（Smart 命令）
		/// </summary>
		public const string ObjectsAlignBottomSmart = "ObjectsAlignBottomSmart";

		/// <summary>
		/// 形状水平平均分布
		/// </summary>
		public const string AlignDistributeHorizontally = "AlignDistributeHorizontally";

		/// <summary>
		/// 形状垂直平均分布
		/// </summary>
		public const string AlignDistributeVertically = "AlignDistributeVertically";

		#endregion

		#region PowerPoint 特定命令 ID

		/// <summary>
		/// 另存为命令 ID
		/// </summary>
		public const int PpFileSaveAs = 101;

		/// <summary>
		/// 从开始放映命令 ID
		/// </summary>
		public const int PpSlideShowFromBeginning = 100;

		/// <summary>
		/// 从当前幻灯片放映命令 ID
		/// </summary>
		public const int PpSlideShowFromCurrent = 101;

		#endregion
	}
}


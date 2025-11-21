using PPA.Core;
using PPA.Manipulation;
using PPA.Utilities;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace PPA.UI.Forms
{
	/// <summary>
	/// 参数设置窗口 用于编辑 PPAConfig.xml 文件（树形结构编辑器）
	/// </summary>
	public partial class SettingsForm:Form
	{
		#region Private Fields

		private TreeView _configTreeView;
		private TextBox _valueTextBox;
		private Button _btnSave;
		private Button _btnCancel;
		private Button _btnReload;
		private Label _pathLabel;
		private SplitContainer _mainSplitContainer;
		private Label _valueLabel;
		private Label _descriptionLabel;
		private TextBox _descriptionTextBox;
		private TableLayoutPanel _rightTableLayout;
		private readonly string _configFilePath;
		private readonly bool _isDesignMode;
		private XmlDocument _xmlDoc;

		#endregion Private Fields

		#region Constructor

		public SettingsForm()
		{
			_isDesignMode=IsDesignMode();
			InitializeComponent();

			if(!_isDesignMode)
			{
				// 运行时再计算路径和本地化，避免设计器解析失败
				string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
				if(string.IsNullOrEmpty(appDataDir))
				{
					appDataDir=Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)??".";
				}
				string ppaConfigDir = Path.Combine(appDataDir, "PPA");
				_configFilePath=Path.Combine(ppaConfigDir,"PPAConfig.xml");

				ApplyLocalization();
				LoadConfigFile();
			}
		}

		#endregion Constructor

		#region Private Methods

		private static bool IsDesignMode()
		{
			try { return LicenseManager.UsageMode==LicenseUsageMode.Designtime; } catch { return false; }
		}

		private void InitializeComponent()
		{
			this.SuspendLayout();

			// --- 窗体基本设置 ---
			this.Text="格式化参数设置";
			this.Size=new Size(1200,700);
			this.StartPosition=FormStartPosition.CenterScreen;
			this.FormBorderStyle=FormBorderStyle.Sizable;
			this.MinimumSize=new Size(1000,500);
			this.MaximumSize=new Size(1400,800);

			// --- 主布局：TableLayoutPanel（1列3行）---
			var mainTableLayout = new TableLayoutPanel
			{
				Dock = DockStyle.Fill,
				ColumnCount = 1,
				RowCount = 3,
				Padding = new Padding(0)
			};
			// 第1行：路径标签（固定高度）
			mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute,40f));
			// 第2行：分割容器（填充剩余空间）
			mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Percent,100f));
			// 第3行：按钮面板（固定高度）
			mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute,60f));

			// --- 1. 第1行：配置文件路径标签 ---
			_pathLabel=new Label
			{
				Text=ResourceManager.GetString("SettingsForm_ConfigFilePath","配置文件路径: ")+(_isDesignMode ? ResourceManager.GetString("SettingsForm_DesignModeConfigFilePath","设计时配置文件路径") : string.Empty),
				Height=30,
				Padding=new Padding(10,10,0,0),
				AutoSize=false,
				Dock=DockStyle.Fill
			};
			mainTableLayout.Controls.Add(_pathLabel,0,0);

			// --- 2. 第3行：按钮面板 ---
			var buttonPanel = new TableLayoutPanel
			{
				Height = 60,
				Padding = new Padding(10),
				ColumnCount = 3,
				RowCount = 1,
				Dock = DockStyle.Fill
			};
			buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100f));
			buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,110f));
			buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,110f));

			_btnReload=new Button
			{
				Text=ResourceManager.GetString("SettingsForm_Reload","重新加载"),
				UseVisualStyleBackColor=true,
				Size=new Size(120,35),
				Font=new Font("Microsoft Sans Serif",10)
			};
			buttonPanel.Controls.Add(_btnReload,0,0);

			_btnCancel=new Button
			{
				Text=ResourceManager.GetString("SettingsForm_Cancel","取消"),
				UseVisualStyleBackColor=true,
				Anchor=AnchorStyles.Right,
				Size=new Size(120,35),
				Font=new Font("Microsoft Sans Serif",10)
			};
			buttonPanel.Controls.Add(_btnCancel,1,0);

			_btnSave=new Button
			{
				Text=ResourceManager.GetString("SettingsForm_Save","保存"),
				UseVisualStyleBackColor=true,
				Anchor=AnchorStyles.Right,
				Size=new Size(120,35),
				Font=new Font("Microsoft Sans Serif",10)
			};
			buttonPanel.Controls.Add(_btnSave,2,0);

			mainTableLayout.Controls.Add(buttonPanel,0,2);

			// --- 3. 第2行：中间主体：分割容器 ---
			_mainSplitContainer=new SplitContainer
			{
				Dock=DockStyle.Fill,
				Orientation=Orientation.Vertical,
				SplitterWidth=8,
				FixedPanel=FixedPanel.None
			};

			// --- 3.1 左侧：树形视图（60%宽度）---
			_configTreeView=new TreeView
			{
				Dock=DockStyle.Fill,
				Font=new Font("Consolas",10),
				HideSelection=false,
				ShowLines=true,
				ShowPlusMinus=true,
				ShowRootLines=true,
				Indent=20
			};
			_configTreeView.AfterSelect+=ConfigTreeView_AfterSelect;
			_configTreeView.BeforeLabelEdit+=ConfigTreeView_BeforeLabelEdit;
			_mainSplitContainer.Panel1.Controls.Add(_configTreeView);

			// --- 3.2 右侧：纵向排列的编辑区和说明区（40%宽度）---
			_rightTableLayout=new TableLayoutPanel
			{
				Dock=DockStyle.Fill,
				ColumnCount=1,
				RowCount=4,
				Padding=new Padding(8,5,8,5)
			};
			// 行0：编辑值标签（固定高度）
			_rightTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute,25f));
			// 行1：编辑值文本框（占编辑区剩余空间，约40%中的大部分）
			_rightTableLayout.RowStyles.Add(new RowStyle(SizeType.Percent,40f));
			// 行2：节点说明标签（固定高度）
			_rightTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute,25f));
			// 行3：节点说明文本框（占说明区剩余空间，约60%中的大部分）
			_rightTableLayout.RowStyles.Add(new RowStyle(SizeType.Percent,60f));

			// 编辑值标签
			_valueLabel=new Label
			{
				Text=ResourceManager.GetString("SettingsForm_EditValue","编辑值:"),
				AutoSize=false,
				Height=25,
				Font=new Font("Microsoft Sans Serif",9,FontStyle.Bold),
				TextAlign=ContentAlignment.MiddleLeft,
				Padding=new Padding(0,3,0,0),
				Dock=DockStyle.Fill
			};
			_rightTableLayout.Controls.Add(_valueLabel,0,0);

			// 编辑值文本框
			_valueTextBox=new TextBox
			{
				Multiline=true,
				Padding=new Padding(8,5,8,5),
				Dock=DockStyle.Fill,
				Font=new Font("Consolas",10),
				AcceptsReturn=true,
				AcceptsTab=true,
				WordWrap=true,
				BorderStyle=BorderStyle.FixedSingle
			};
			_valueTextBox.TextChanged+=ValueTextBox_TextChanged;
			_rightTableLayout.Controls.Add(_valueTextBox,0,1);

			// 节点说明标签
			_descriptionLabel=new Label
			{
				Text=ResourceManager.GetString("SettingsForm_NodeDescription","节点说明:"),
				AutoSize=false,
				Height=25,
				Font=new Font("Microsoft Sans Serif",9,FontStyle.Bold),
				TextAlign=ContentAlignment.MiddleLeft,
				Padding=new Padding(0,3,0,0),
				Dock=DockStyle.Fill
			};
			_rightTableLayout.Controls.Add(_descriptionLabel,0,2);

			// 节点说明文本框
			_descriptionTextBox=new TextBox
			{
				Multiline=true,
				Dock=DockStyle.Fill,
				Font=new Font("Microsoft Sans Serif",9),
				ScrollBars=ScrollBars.Vertical,
				ReadOnly=true,
				BackColor=SystemColors.ControlLight,
				BorderStyle=BorderStyle.FixedSingle,
				WordWrap=true
			};
			_rightTableLayout.Controls.Add(_descriptionTextBox,0,3);

			_mainSplitContainer.Panel2.Controls.Add(_rightTableLayout);
			mainTableLayout.Controls.Add(_mainSplitContainer,0,1);

			// 将主布局添加到窗体
			this.Controls.Add(mainTableLayout);

			// --- 事件绑定 ---
			_btnReload.Click+=BtnReload_Click;
			_btnSave.Click+=BtnSave_Click;
			_btnCancel.Click+=BtnCancel_Click;

			// 在窗体加载后设置分割距离，确保使用正确的客户端宽度
			this.Load+=(s,e) =>
			{
				UpdateSplitterDistance();
			};

			// 窗体大小改变时保持左右 60/40 比例
			this.SizeChanged+=(s,e) =>
			{
				UpdateSplitterDistance();
			};

			this.ResumeLayout(false);
		}

		// 新增方法：更新分割器距离
		private void UpdateSplitterDistance()
		{
			try
			{
				_mainSplitContainer.SplitterDistance=(int) (_mainSplitContainer.Width*0.6);
			} catch
			{
				// 忽略可能的异常
			}
		}

		private void ApplyLocalization()
		{
			this.Text=ResourceManager.GetString("SettingsForm_Title","格式化参数设置");
			_btnReload.Text=ResourceManager.GetString("SettingsForm_Reload","重新加载");
			_btnSave.Text=ResourceManager.GetString("SettingsForm_Save","保存");
			_btnCancel.Text=ResourceManager.GetString("SettingsForm_Cancel","取消");
			_pathLabel.Text=ResourceManager.GetString("SettingsForm_ConfigFilePath","配置文件路径: ")+_configFilePath;
			_valueLabel.Text=ResourceManager.GetString("SettingsForm_EditValue","编辑值:");
			_descriptionLabel.Text=ResourceManager.GetString("SettingsForm_NodeDescription","节点说明:");
		}

		private void LoadConfigFile()
		{
			try
			{
				_xmlDoc=new XmlDocument();

				if(File.Exists(_configFilePath))
				{
					_xmlDoc.Load(_configFilePath);
				} else
				{
					var defaultConfig = new FormattingConfig();
					defaultConfig.Save();
					_xmlDoc.Load(_configFilePath);
				}

				BuildTreeFromXml();
			} catch(Exception ex)
			{
				MessageBox.Show(
					ResourceManager.GetString("SettingsForm_LoadError",ex.Message,"加载配置文件失败: {0}"),
					ResourceManager.GetString("SettingsForm_Error","错误"),
					MessageBoxButtons.OK,MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// 从 XML 文档构建树形结构
		/// </summary>
		private void BuildTreeFromXml()
		{
			_configTreeView.BeginUpdate();
			_configTreeView.Nodes.Clear();

			if(_xmlDoc?.DocumentElement!=null)
			{
				TreeNode rootNode = CreateTreeNode(_xmlDoc.DocumentElement);
				_configTreeView.Nodes.Add(rootNode);
				rootNode.Expand();
			}

			_configTreeView.EndUpdate();
		}

		/// <summary>
		/// 递归创建树节点
		/// </summary>
		private TreeNode CreateTreeNode(XmlNode xmlNode)
		{
			string nodeText = xmlNode.Name;

			if(xmlNode.NodeType==XmlNodeType.Attribute)
			{
				nodeText=$"{xmlNode.Name} = {xmlNode.Value}";
			} else if(xmlNode.NodeType==XmlNodeType.Element)
			{
				if(xmlNode.HasChildNodes&&xmlNode.FirstChild.NodeType==XmlNodeType.Element)
				{
					nodeText=xmlNode.Name;
				} else if(xmlNode.HasChildNodes&&xmlNode.FirstChild.NodeType==XmlNodeType.Text)
				{
					nodeText=$"{xmlNode.Name} = {xmlNode.InnerText}";
				} else if(!xmlNode.HasChildNodes&&xmlNode.Attributes!=null&&xmlNode.Attributes.Count>0)
				{
					nodeText=xmlNode.Name;
				}
			}

			TreeNode treeNode = new(nodeText)
			{
				Tag = xmlNode
			};

			if(xmlNode.Attributes!=null)
			{
				foreach(XmlAttribute attr in xmlNode.Attributes)
				{
					TreeNode attrNode = new($"{attr.Name} = {attr.Value}")
					{
						Tag = attr,
						ForeColor = Color.Blue
					};
					treeNode.Nodes.Add(attrNode);
				}
			}

			foreach(XmlNode childNode in xmlNode.ChildNodes)
			{
				if(childNode.NodeType==XmlNodeType.Element)
				{
					TreeNode childTreeNode = CreateTreeNode(childNode);
					treeNode.Nodes.Add(childTreeNode);
				} else if(childNode.NodeType==XmlNodeType.Text&&string.IsNullOrWhiteSpace(childNode.Value))
				{
					continue;
				}
			}

			return treeNode;
		}

		/// <summary>
		/// 树节点选择事件
		/// </summary>
		private void ConfigTreeView_AfterSelect(object sender,TreeViewEventArgs e)
		{
			if(e.Node?.Tag is XmlNode xmlNode)
			{
				UpdateValueEditor(xmlNode);
				UpdateNodeDescription(xmlNode);
			} else
			{
				_valueTextBox.Text=string.Empty;
				_valueTextBox.ReadOnly=true;
				_descriptionTextBox.Text=ResourceManager.GetString("ConfigDesc_NoNodeSelected","未选择有效节点");
			}
		}

		/// <summary>
		/// 更新值编辑器
		/// </summary>
		private void UpdateValueEditor(XmlNode xmlNode)
		{
			_valueTextBox.TextChanged-=ValueTextBox_TextChanged;

			try
			{
				if(xmlNode.NodeType==XmlNodeType.Attribute)
				{
					_valueTextBox.Text=xmlNode.Value??string.Empty;
					_valueTextBox.ReadOnly=false;
					_valueTextBox.BackColor=SystemColors.Window;
				} else if(xmlNode.NodeType==XmlNodeType.Element)
				{
					bool hasElementChildren = false;
					bool hasTextContent = false;
					string textContent = string.Empty;

					foreach(XmlNode child in xmlNode.ChildNodes)
					{
						if(child.NodeType==XmlNodeType.Element)
						{
							hasElementChildren=true;
							break;
						} else if(child.NodeType==XmlNodeType.Text&&!string.IsNullOrWhiteSpace(child.Value))
						{
							hasTextContent=true;
							textContent=child.Value;
						}
					}

					if(hasTextContent&&!hasElementChildren)
					{
						_valueTextBox.Text=textContent;
						_valueTextBox.ReadOnly=false;
						_valueTextBox.BackColor=SystemColors.Window;
					} else if(!hasElementChildren&&!hasTextContent&&xmlNode.Attributes!=null&&xmlNode.Attributes.Count>0)
					{
						if(xmlNode.Attributes.Count==1)
						{
							_valueTextBox.Text=xmlNode.Attributes[0].Value??string.Empty;
							_valueTextBox.ReadOnly=false;
							_valueTextBox.BackColor=SystemColors.Window;
						} else
						{
							_valueTextBox.Text=ResourceManager.GetString("ConfigDesc_NodeHasMultipleAttributes",xmlNode.Attributes.Count,"节点包含 {0} 个属性");
							_valueTextBox.ReadOnly=true;
							_valueTextBox.BackColor=SystemColors.Control;
						}
					} else
					{
						_valueTextBox.Text=ResourceManager.GetString("ConfigDesc_NodeInfo",xmlNode.Name,xmlNode.ChildNodes.Count,xmlNode.Attributes?.Count??0,"节点: {0}\r\n子节点数: {1}\r\n属性数: {2}");
						_valueTextBox.ReadOnly=true;
						_valueTextBox.BackColor=SystemColors.Control;
					}
				}

				// 确保光标可见
				_valueTextBox.SelectionStart=0;
				_valueTextBox.SelectionLength=0;
				_valueTextBox.ScrollToCaret();
			} finally
			{
				_valueTextBox.TextChanged+=ValueTextBox_TextChanged;
			}
		}

		/// <summary>
		/// 更新节点说明
		/// </summary>
		private void UpdateNodeDescription(XmlNode xmlNode)
		{
			string description = GenerateNodeDescription(xmlNode);
			_descriptionTextBox.Text=description;
		}

		/// <summary>
		/// 生成节点说明文本
		/// </summary>
		private string GenerateNodeDescription(XmlNode xmlNode)
		{
			var description = new System.Text.StringBuilder();
			string nodePath = GetNodeDisplayPath(xmlNode);

			if(xmlNode.NodeType==XmlNodeType.Attribute)
			{
				description.AppendLine($"{ResourceManager.GetString("ConfigDesc_Label_Attribute","● 属性:")} {xmlNode.Name}");
				description.AppendLine($"{ResourceManager.GetString("ConfigDesc_Label_Path","● 路径:")} {nodePath}");
				description.AppendLine($"{ResourceManager.GetString("ConfigDesc_Label_CurrentValue","● 当前值:")} {xmlNode.Value}");
				description.AppendLine();
				var specificAttrDesc = GetAttributeDescriptionByPath(nodePath);
				description.AppendLine(!string.IsNullOrEmpty(specificAttrDesc) ? specificAttrDesc : ResourceManager.GetString("ConfigDesc_NoSpecificDescription","暂无针对性说明。"));
			} else if(xmlNode.NodeType==XmlNodeType.Element)
			{
				description.AppendLine($"{ResourceManager.GetString("ConfigDesc_Label_Element","● 元素:")} {xmlNode.Name}");
				description.AppendLine($"{ResourceManager.GetString("ConfigDesc_Label_Path","● 路径:")} {nodePath}");
				description.AppendLine($"{ResourceManager.GetString("ConfigDesc_Label_NodeType","● 节点类型:")} {GetNodeTypeDescription(xmlNode)}");
				description.AppendLine($"{ResourceManager.GetString("ConfigDesc_Label_ChildCount","● 子节点数:")} {xmlNode.ChildNodes.Count}");
				description.AppendLine($"{ResourceManager.GetString("ConfigDesc_Label_AttributeCount","● 属性数:")} {xmlNode.Attributes?.Count??0}");
				description.AppendLine();
				var specificElemDesc = GetElementDescriptionByPath(nodePath);
				description.AppendLine(!string.IsNullOrEmpty(specificElemDesc) ? specificElemDesc : ResourceManager.GetString("ConfigDesc_NoSpecificDescription","暂无针对性说明。"));

				// 添加属性信息
				if(xmlNode.Attributes!=null&&xmlNode.Attributes.Count>0)
				{
					description.AppendLine();
					description.AppendLine(ResourceManager.GetString("ConfigDesc_Label_ContainsAttributes","包含属性:"));
					foreach(XmlAttribute attr in xmlNode.Attributes)
					{
						description.AppendLine($"  - {attr.Name} = {attr.Value}");
					}
				}
			}

			return description.ToString();
		}

		/// <summary>
		/// 生成节点显示路径（简化的 XPath）
		/// </summary>
		private string GetNodeDisplayPath(XmlNode node)
		{
			if(node==null) return "/";
			var parts = new System.Collections.Generic.List<string>();
			XmlNode current = node.NodeType == XmlNodeType.Attribute ? ((XmlAttribute)node).OwnerElement : node;
			while(current!=null&&current.NodeType==XmlNodeType.Element)
			{
				string name = current.Name;
				int index = 1;
				if(current.ParentNode!=null)
				{
					foreach(XmlNode sibling in current.ParentNode.ChildNodes)
					{
						if(sibling==current) break;
						if(sibling.NodeType==XmlNodeType.Element&&sibling.Name==name) index++;
					}
				}
				parts.Add(index>1 ? $"{name}[{index}]" : name);
				current=current.ParentNode;
			}
			parts.Reverse();
			string path = "/" + string.Join("/", parts);
			if(node.NodeType==XmlNodeType.Attribute)
			{
				path+=$"/@{node.Name}";
			}
			return path;
		}

		/// <summary>
		/// 按路径提供精确的属性说明（与配置结构一一对应）
		/// </summary>
		private string GetAttributeDescriptionByPath(string path)
		{
			string p = (path ?? string.Empty).Trim();

			// TableFormattingConfig
			if(p.EndsWith("/PPAConfig/Table/@StyleId",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Table_StyleId","表格样式 ID（Office 主题样式的 GUID）。用于应用内置主题表格样式。");
			if(p.EndsWith("/PPAConfig/Table/@DataRowBorderWidth",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Table_DataRowBorderWidth","数据行边框宽度（磅）。");
			if(p.EndsWith("/PPAConfig/Table/@HeaderRowBorderWidth",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Table_HeaderRowBorderWidth","标题行边框宽度（磅）。");
			if(p.EndsWith("/PPAConfig/Table/@DataRowBorderColor",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Table_DataRowBorderColor","数据行边框颜色主题（如 Accent1~Accent6、Dark1、Light1）。");
			if(p.EndsWith("/PPAConfig/Table/@HeaderRowBorderColor",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Table_HeaderRowBorderColor","标题行边框颜色主题（如 Accent1~Accent6、Dark1、Light1）。");
			if(p.EndsWith("/PPAConfig/Table/@AutoNumberFormat",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Table_AutoNumberFormat","是否启用数字格式化（true/false）。");
			if(p.EndsWith("/PPAConfig/Table/@DecimalPlaces",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Table_DecimalPlaces","数字格式化保留的小数位数。");
			if(p.EndsWith("/PPAConfig/Table/@NegativeTextColor",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Table_NegativeTextColor","负数文本颜色（OLE 整数 RGB，255 表示红色）。");

			// TableSettingsConfig
			if(p.EndsWith("/PPAConfig/Table/TableSettings/@FirstRow",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_TableSettings_FirstRow","启用“标题行”格式（true/false）。");
			if(p.EndsWith("/PPAConfig/Table/TableSettings/@FirstCol",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_TableSettings_FirstCol","启用“首列”格式（true/false）。");
			if(p.EndsWith("/PPAConfig/Table/TableSettings/@LastRow",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_TableSettings_LastRow","启用“汇总行/最后一行”格式（true/false）。");
			if(p.EndsWith("/PPAConfig/Table/TableSettings/@LastCol",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_TableSettings_LastCol","启用“末列”格式（true/false）。");
			if(p.EndsWith("/PPAConfig/Table/TableSettings/@HorizBanding",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_TableSettings_HorizBanding","启用“交错着色（横向条纹）”（true/false）。");
			if(p.EndsWith("/PPAConfig/Table/TableSettings/@VertBanding",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_TableSettings_VertBanding","启用“交错着色（纵向条纹）”（true/false）。");

			// TextFormattingConfig
			if(p.EndsWith("/PPAConfig/Text/@LeftIndent",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_LeftIndent","段落左缩进（厘米）。");

			// MarginsConfig
			if(p.EndsWith("/PPAConfig/Text/Margins/@Top",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Margins_Top","文本框上边距（厘米）。");
			if(p.EndsWith("/PPAConfig/Text/Margins/@Bottom",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Margins_Bottom","文本框下边距（厘米）。");
			if(p.EndsWith("/PPAConfig/Text/Margins/@Left",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Margins_Left","文本框左边距（厘米）。");
			if(p.EndsWith("/PPAConfig/Text/Margins/@Right",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Margins_Right","文本框右边距（厘米）。");

			// FontConfig（Table.DataRowFont / Table.HeaderRowFont / Text.Font / Chart.RegularFont / Chart.TitleFont）
			if(p.EndsWith("/@Name",StringComparison.OrdinalIgnoreCase)&&p.IndexOf("/Font",StringComparison.OrdinalIgnoreCase)>=0)
				return ResourceManager.GetString("ConfigDesc_Attr_Font_Name","字体名称（西文字体），如 Arial、Consolas。");
			if(p.EndsWith("/@NameFarEast",StringComparison.OrdinalIgnoreCase)&&p.IndexOf("/Font",StringComparison.OrdinalIgnoreCase)>=0)
				return ResourceManager.GetString("ConfigDesc_Attr_Font_NameFarEast","中文/东亚字体名称（如 +mn-ea）。");
			if(p.EndsWith("/@Size",StringComparison.OrdinalIgnoreCase)&&p.IndexOf("/Font",StringComparison.OrdinalIgnoreCase)>=0)
				return ResourceManager.GetString("ConfigDesc_Attr_Font_Size","字体大小（磅）。");
			if(p.EndsWith("/@Bold",StringComparison.OrdinalIgnoreCase)&&p.IndexOf("/Font",StringComparison.OrdinalIgnoreCase)>=0)
				return ResourceManager.GetString("ConfigDesc_Attr_Font_Bold","是否加粗（true/false）。");
			if(p.EndsWith("/@ThemeColor",StringComparison.OrdinalIgnoreCase)&&p.IndexOf("/Font",StringComparison.OrdinalIgnoreCase)>=0)
				return ResourceManager.GetString("ConfigDesc_Attr_Font_ThemeColor","字体主题颜色名称（如 Dark1、Accent1~Accent6）。");

			// ParagraphConfig
			if(p.EndsWith("/PPAConfig/Text/Paragraph/@Alignment",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Paragraph_Alignment","段落对齐方式：Left、Center、Right、Justify、Distribute。");
			if(p.EndsWith("/PPAConfig/Text/Paragraph/@WordWrap",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Paragraph_WordWrap","是否自动换行（true/false）。");
			if(p.EndsWith("/PPAConfig/Text/Paragraph/@SpaceBefore",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Paragraph_SpaceBefore","段前间距（磅）。");
			if(p.EndsWith("/PPAConfig/Text/Paragraph/@SpaceAfter",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Paragraph_SpaceAfter","段后间距（磅）。");
			if(p.EndsWith("/PPAConfig/Text/Paragraph/@SpaceWithin",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Paragraph_SpaceWithin","行距（倍数）。");
			if(p.EndsWith("/PPAConfig/Text/Paragraph/@FarEastLineBreakControl",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Paragraph_FarEastLineBreakControl","中文断行控制（true/false）。");
			if(p.EndsWith("/PPAConfig/Text/Paragraph/@HangingPunctuation",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Paragraph_HangingPunctuation","标点悬挂（true/false）。");

			// BulletConfig
			if(p.EndsWith("/PPAConfig/Text/Bullet/@Type",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Bullet_Type","项目符号类型：None、Numbered、Unnumbered、Picture。");
			if(p.EndsWith("/PPAConfig/Text/Bullet/@Character",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Bullet_Character","项目符号字符（Unicode 码点，如 9632 为实心方块）。");
			if(p.EndsWith("/PPAConfig/Text/Bullet/@FontName",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Bullet_FontName","项目符号字体名称（用于正确显示符号）。");
			if(p.EndsWith("/PPAConfig/Text/Bullet/@RelativeSize",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Bullet_RelativeSize","项目符号相对字号（相对于段落字体）。");
			if(p.EndsWith("/PPAConfig/Text/Bullet/@ThemeColor",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Text_Bullet_ThemeColor","项目符号主题颜色名称。");

			// ShortcutsConfig
			if(p.EndsWith("/PPAConfig/Shortcuts/@FormatTables",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Shortcuts_FormatTables","“美化表格”快捷键：填写数字或字母（如 1、T），实际为 Ctrl+键。");
			if(p.EndsWith("/PPAConfig/Shortcuts/@FormatText",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Shortcuts_FormatText","“美化文本”快捷键：填写数字或字母（如 2、X），实际为 Ctrl+键。");
			if(p.EndsWith("/PPAConfig/Shortcuts/@FormatChart",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Shortcuts_FormatChart","“美化图表”快捷键：填写数字或字母（如 3、C），实际为 Ctrl+键。");
			if(p.EndsWith("/PPAConfig/Shortcuts/@CreateBoundingBox",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Shortcuts_CreateBoundingBox","“插入形状”快捷键：填写数字或字母（如 4、I），实际为 Ctrl+键。");
			
			// LoggingConfig
			if(p.EndsWith("/PPAConfig/Logging/@EnableFileLogging",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Logging_EnableFileLogging","是否启用文件日志记录（true/false）。");
			if(p.EndsWith("/PPAConfig/Logging/@MaxLogFiles",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Logging_MaxLogFiles","最多保留的日志文件数量。");
			if(p.EndsWith("/PPAConfig/Logging/@MaxLogAgeDays",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Logging_MaxLogAgeDays","日志文件最长保留天数（天）。");
			if(p.EndsWith("/PPAConfig/Logging/@MinimumLogLevel",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Attr_Logging_MinimumLogLevel","最小写入日志级别（Debug、Information、Warning、Error）。");

			return string.Empty;
		}

		/// <summary>
		/// 按路径提供精确的元素说明（与配置结构一一对应）
		/// </summary>
		private string GetElementDescriptionByPath(string path)
		{
			string p = (path ?? string.Empty).Trim();

			if(p.Equals("/PPAConfig",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_PPAConfig","PPA 配置根节点，包含表格、文本、图表格式及快捷键配置。");
			if(p.EndsWith("/PPAConfig/Table",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Table","表格格式化配置：样式、边框、数字格式及表格标志设置。");
			if(p.EndsWith("/PPAConfig/Table/DataRowFont",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Table_DataRowFont","数据行字体配置：名称、大小、粗细、主题颜色。");
			if(p.EndsWith("/PPAConfig/Table/HeaderRowFont",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Table_HeaderRowFont","标题行字体配置：名称、大小、粗细、主题颜色。");
			if(p.EndsWith("/PPAConfig/Table/TableSettings",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Table_TableSettings","表格标志设置：标题行、首列、末列、汇总行、条纹等。");

			if(p.EndsWith("/PPAConfig/Text",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Text","文本格式化配置：边距、字体、段落、项目符号等。");
			if(p.EndsWith("/PPAConfig/Text/Margins",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Text_Margins","文本边距（厘米）：上/下/左/右。");
			if(p.EndsWith("/PPAConfig/Text/Font",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Text_Font","文本字体配置：名称、中文字体、大小、粗细、主题颜色。");
			if(p.EndsWith("/PPAConfig/Text/Paragraph",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Text_Paragraph","段落格式：对齐、换行、段前/段后、行距、中文断行、标点悬挂。");
			if(p.EndsWith("/PPAConfig/Text/Bullet",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Text_Bullet","项目符号：类型、符号字符、字体、相对字号、颜色。");

			if(p.EndsWith("/PPAConfig/Chart",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Chart","图表格式化配置：常规与标题的字体样式。");
			if(p.EndsWith("/PPAConfig/Chart/RegularFont",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Chart_RegularFont","图表常规字体：名称、中文字体、大小、粗细、主题颜色。");
			if(p.EndsWith("/PPAConfig/Chart/TitleFont",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Chart_TitleFont","图表标题字体：名称、中文字体、大小、粗细、主题颜色。");

			if(p.EndsWith("/PPAConfig/Shortcuts",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Shortcuts","快捷键配置：填写数字或字母，实际组合为 Ctrl+键。");
			if(p.EndsWith("/PPAConfig/Logging",StringComparison.OrdinalIgnoreCase))
				return ResourceManager.GetString("ConfigDesc_Elem_Logging","日志配置：文件日志开关、日志保留数量/天数、最小日志级别。");

			return string.Empty;
		}

		/// <summary>
		/// 获取节点类型描述
		/// </summary>
		private string GetNodeTypeDescription(XmlNode xmlNode)
		{
			if(xmlNode.HasChildNodes)
			{
				bool hasElements = false;
				bool hasText = false;

				foreach(XmlNode child in xmlNode.ChildNodes)
				{
					if(child.NodeType==XmlNodeType.Element) hasElements=true;
					if(child.NodeType==XmlNodeType.Text&&!string.IsNullOrWhiteSpace(child.Value)) hasText=true;
				}

				if(hasElements&&hasText) return ResourceManager.GetString("ConfigDesc_NodeType_Mixed","混合节点");
				if(hasElements) return ResourceManager.GetString("ConfigDesc_NodeType_Container","容器节点");
				if(hasText) return ResourceManager.GetString("ConfigDesc_NodeType_Text","文本节点");
			}

			if(xmlNode.Attributes!=null&&xmlNode.Attributes.Count>0)
				return ResourceManager.GetString("ConfigDesc_NodeType_Attribute","属性节点");

			return ResourceManager.GetString("ConfigDesc_NodeType_Empty","空节点");
		}

		/// <summary>
		/// 值文本框内容改变事件
		/// </summary>
		private void ValueTextBox_TextChanged(object sender,EventArgs e)
		{
			if(_configTreeView.SelectedNode?.Tag is XmlNode xmlNode&&!_valueTextBox.ReadOnly)
			{
				if(xmlNode.NodeType==XmlNodeType.Attribute)
				{
					((XmlAttribute) xmlNode).Value=_valueTextBox.Text;
					_configTreeView.SelectedNode.Text=$"{xmlNode.Name} = {_valueTextBox.Text}";

					// 更新说明
					UpdateNodeDescription(xmlNode);
				} else if(xmlNode.NodeType==XmlNodeType.Element)
				{
					xmlNode.InnerText=_valueTextBox.Text;
					_configTreeView.SelectedNode.Text=$"{xmlNode.Name} = {_valueTextBox.Text}";

					// 更新说明
					UpdateNodeDescription(xmlNode);
				}
			}
		}

		/// <summary>
		/// 树节点标签编辑前事件（禁止编辑节点名称）
		/// </summary>
		private void ConfigTreeView_BeforeLabelEdit(object sender,NodeLabelEditEventArgs e)
		{
			e.CancelEdit=true;
		}

		private void BtnReload_Click(object sender,EventArgs e)
		{
			var result = MessageBox.Show(
				ResourceManager.GetString("SettingsForm_ReloadConfirm", "重新加载将丢失当前未保存的修改，是否继续？"),
				ResourceManager.GetString("SettingsForm_Confirm", "确认"),
				MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			if(result==DialogResult.Yes)
			{
				LoadConfigFile();
				FormattingConfig.Reload();
				KeyboardShortcutHelper.ReloadShortcuts();
				Toast.Show(ResourceManager.GetString("Settings_ConfigReloaded","配置已重新加载"),Toast.ToastType.Success);
			}
		}

		private void BtnSave_Click(object sender,EventArgs e)
		{
			try
			{
				if(_xmlDoc==null)
				{
					throw new InvalidOperationException("XML 文档未加载");
				}

				_xmlDoc.Save(_configFilePath);

				FormattingConfig.Reload();
				KeyboardShortcutHelper.ReloadShortcuts();
				Toast.Show(ResourceManager.GetString("Settings_ConfigSaved","配置已保存"),Toast.ToastType.Success);
				this.DialogResult=DialogResult.OK;
				this.Close();
			} catch(Exception ex)
			{
				MessageBox.Show(
					ResourceManager.GetString("SettingsForm_SaveError",ex.Message,"保存配置文件失败: {0}"),
					ResourceManager.GetString("SettingsForm_Error","错误"),
					MessageBoxButtons.OK,MessageBoxIcon.Error);
			}
		}

		private void BtnCancel_Click(object sender,EventArgs e)
		{
			var result = MessageBox.Show(
				ResourceManager.GetString("SettingsForm_DiscardConfirm", "是否放弃当前修改？"),
				ResourceManager.GetString("SettingsForm_Confirm", "确认"),
				MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			if(result==DialogResult.Yes)
			{
				this.DialogResult=DialogResult.Cancel;
				this.Close();
			}
		}

		#endregion Private Methods
	}
}

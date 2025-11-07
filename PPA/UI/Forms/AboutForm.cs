using PPA.Core;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace PPA.UI.Forms
{
	/// <summary>
	/// 关于窗口 显示插件版本信息和说明
	/// </summary>
	public partial class AboutForm:Form
	{
		#region Private Fields

		private Label _titleLabel;
		private Label _versionLabel;
		private Label _descriptionLabel;
		private Label _copyrightLabel;
		private Button _btnClose;
		private LinkLabel _linkLabel;

		#endregion Private Fields

		#region Constructor

		public AboutForm()
		{
			InitializeComponent();
			LoadVersionInfo();
		}

		#endregion Constructor

		#region Private Methods

		private void InitializeComponent()
		{
			this.Text=ResourceManager.GetString("AboutForm_Title","关于 PPA");
			this.Size=new System.Drawing.Size(500,420);
			this.StartPosition=FormStartPosition.CenterScreen;
			this.FormBorderStyle=FormBorderStyle.FixedDialog;
			this.MaximizeBox=false;
			this.MinimizeBox=false;
			this.Padding=new Padding(20);

			// 创建主面板，使用 TableLayoutPanel 实现更好的布局
			var mainPanel = new TableLayoutPanel
			{
				Dock = DockStyle.Fill,
				ColumnCount = 1,
				RowCount = 5,
				AutoSize = true
			};
			mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100F));
			mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
			mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
			mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
			mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
			mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
			this.Controls.Add(mainPanel);

			// 标题
			_titleLabel=new Label
			{
				Text=ResourceManager.GetString("AboutForm_TitleText","PPA - PowerPoint Assistant"),
				Font=new System.Drawing.Font("Segoe UI",18,System.Drawing.FontStyle.Bold),
				AutoSize=true,
				TextAlign=System.Drawing.ContentAlignment.MiddleCenter,
				Dock=DockStyle.Fill,
				Margin=new Padding(0,0,0,10)
			};
			mainPanel.Controls.Add(_titleLabel,0,0);

			// 版本信息
			_versionLabel=new Label
			{
				AutoSize=true,
				Font=new System.Drawing.Font("Segoe UI",10),
				TextAlign=System.Drawing.ContentAlignment.MiddleCenter,
				Dock=DockStyle.Fill,
				Margin=new Padding(0,0,0,15)
			};
			mainPanel.Controls.Add(_versionLabel,0,1);

			// 描述
			_descriptionLabel=new Label
			{
				Text=ResourceManager.GetString("AboutForm_Description","PowerPoint 插件，提供表格、文本、图表格式化功能，以及形状对齐、吸附等辅助工具。\n\n支持多语言界面，可自定义格式化参数。"),
				AutoSize=false,
				Size=new System.Drawing.Size(460,120),
				Font=new System.Drawing.Font("Segoe UI",9.5f),
				TextAlign=System.Drawing.ContentAlignment.MiddleLeft,
				Dock=DockStyle.Fill,
				Margin=new Padding(0,0,0,15)
			};
			mainPanel.Controls.Add(_descriptionLabel,0,2);

			// 版权信息和链接的容器
			var bottomPanel = new FlowLayoutPanel
			{
				Dock = DockStyle.Fill,
				FlowDirection = FlowDirection.LeftToRight,
				WrapContents = false,
				AutoSize = true,
				Margin = new Padding(0, 0, 0, 10)
			};
			mainPanel.Controls.Add(bottomPanel,0,3);

			// 版权信息
			_copyrightLabel=new Label
			{
				AutoSize=true,
				Font=new System.Drawing.Font("Segoe UI",8),
				ForeColor=System.Drawing.Color.Gray,
				Margin=new Padding(0,0,20,0)
			};
			bottomPanel.Controls.Add(_copyrightLabel);

			// 链接标签
			_linkLabel=new LinkLabel
			{
				Text=ResourceManager.GetString("AboutForm_ViewDocs","查看项目文档"),
				AutoSize=true,
				Font=new System.Drawing.Font("Segoe UI",9),
				Margin=new Padding(0)
			};
			_linkLabel.LinkClicked+=LinkLabel_LinkClicked;
			bottomPanel.Controls.Add(_linkLabel);

			// 按钮面板
			var buttonPanel = new Panel
			{
				Dock = DockStyle.Fill,
				Height = 40
			};
			mainPanel.Controls.Add(buttonPanel,0,4);

			// 关闭按钮
			_btnClose=new Button
			{
				Text=ResourceManager.GetString("AboutForm_Close","关闭"),
				Size=new System.Drawing.Size(100,32),
				Location=new System.Drawing.Point(390,4),
				Anchor=AnchorStyles.Bottom|AnchorStyles.Right,
				DialogResult=DialogResult.OK,
				Font=new System.Drawing.Font("Segoe UI",9)
			};
			_btnClose.Click+=(s,e) => this.Close();
			buttonPanel.Controls.Add(_btnClose);
		}

		private void LoadVersionInfo()
		{
			try
			{
				// 从程序集获取版本号（主版本源）
				var assembly = Assembly.GetExecutingAssembly();
				var version = assembly.GetName().Version;

				// 格式化版本号：如果修订号为 0，显示为三段式（如 "0.9.0"），否则显示四段式（如 "0.9.0.1"）
				string versionText;
				if(version!=null)
				{
					if(version.Revision==0)
					{
						versionText=$"{version.Major}.{version.Minor}.{version.Build}";
					} else
					{
						versionText=version.ToString();
					}
				} else
				{
					versionText=ResourceManager.GetString("AboutForm_Unknown","未知");
				}

				_versionLabel.Text=ResourceManager.GetString("AboutForm_Version",versionText,"版本: {0}");

				// 加载版权信息
				try
				{
					var fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);

					if(fileVersionInfo!=null)
					{
						var copyright = fileVersionInfo.LegalCopyright;
						if(!string.IsNullOrEmpty(copyright))
						{
							_copyrightLabel.Text=copyright;
						} else
						{
							_copyrightLabel.Text=$"Copyright © {DateTime.Now.Year}";
						}
					} else
					{
						_copyrightLabel.Text=$"Copyright © {DateTime.Now.Year}";
					}
				} catch
				{
					_copyrightLabel.Text=$"Copyright © {DateTime.Now.Year}";
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"加载版本信息失败: {ex.Message}","WARN");
				_versionLabel.Text=ResourceManager.GetString("AboutForm_Version",ResourceManager.GetString("AboutForm_Unknown","未知"),"版本: {0}");
				_copyrightLabel.Text=$"Copyright © {DateTime.Now.Year}";
			}
		}

		private void LinkLabel_LinkClicked(object sender,LinkLabelLinkClickedEventArgs e)
		{
			try
			{
				Process.Start(new ProcessStartInfo
				{
					FileName="https://github.com/sandorn/PPA",
					UseShellExecute=true
				});
			} catch(Exception ex)
			{
				Profiler.LogMessage($"打开链接失败: {ex.Message}","WARN");
				MessageBox.Show(
					ResourceManager.GetString("AboutForm_LinkError",ex.Message,"无法打开链接: {0}"),
					ResourceManager.GetString("AboutForm_Error","错误"),
					MessageBoxButtons.OK,MessageBoxIcon.Warning);
			}
		}

		#endregion Private Methods
	}
}

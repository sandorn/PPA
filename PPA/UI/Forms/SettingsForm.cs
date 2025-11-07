using PPA.Core;
using PPA.Formatting;
using PPA.Utilities;
using System;
using System.IO;
using System.Windows.Forms;

namespace PPA.UI.Forms
{
	/// <summary>
	/// 参数设置窗口 用于编辑 PPAConfig.xml 文件
	/// </summary>
	public partial class SettingsForm:Form
	{
		#region Private Fields

		private TextBox _configTextBox;
		private Button _btnSave;
		private Button _btnCancel;
		private Button _btnReload;
		private string _configFilePath;

		#endregion Private Fields

		#region Constructor

		public SettingsForm()
		{
			InitializeComponent();
			LoadConfigFile();
		}

		#endregion Constructor

		#region Private Methods

		private void InitializeComponent()
		{
			this.Text=ResourceManager.GetString("SettingsForm_Title","格式化参数设置");
			this.Size=new System.Drawing.Size(800,600);
			this.StartPosition=FormStartPosition.CenterScreen;
			this.FormBorderStyle=FormBorderStyle.Sizable;
			this.MinimumSize=new System.Drawing.Size(600,400);

			// 配置文件路径（使用 AppData 目录）
			string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
			if(string.IsNullOrEmpty(appDataDir))
			{
				appDataDir=Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)??".";
			}
			string ppaConfigDir = Path.Combine(appDataDir, "PPA");
			_configFilePath=Path.Combine(ppaConfigDir,"PPAConfig.xml");

			// 创建文本框
			_configTextBox=new TextBox
			{
				Multiline=true,
				ScrollBars=ScrollBars.Vertical, // 启用自动换行后，只需垂直滚动条
				Dock=DockStyle.Fill,
				Font=new System.Drawing.Font("Consolas",9),
				AcceptsTab=true,
				WordWrap=true // 启用自动换行，避免单行过长
			};
			this.Controls.Add(_configTextBox);

			// 创建按钮面板
			var buttonPanel = new Panel
			{
				Dock = DockStyle.Bottom,
				Height = 50
			};
			this.Controls.Add(buttonPanel);

			// 重新加载按钮
			_btnReload=new Button
			{
				Text=ResourceManager.GetString("SettingsForm_Reload","重新加载"),
				Size=new System.Drawing.Size(100,30),
				Location=new System.Drawing.Point(10,10),
				Anchor=AnchorStyles.Bottom|AnchorStyles.Left
			};
			_btnReload.Click+=BtnReload_Click;
			buttonPanel.Controls.Add(_btnReload);

			// 保存按钮
			_btnSave=new Button
			{
				Text=ResourceManager.GetString("SettingsForm_Save","保存"),
				Size=new System.Drawing.Size(100,30),
				Location=new System.Drawing.Point(580,10),
				Anchor=AnchorStyles.Bottom|AnchorStyles.Right
			};
			_btnSave.Click+=BtnSave_Click;
			buttonPanel.Controls.Add(_btnSave);

			// 取消按钮
			_btnCancel=new Button
			{
				Text=ResourceManager.GetString("SettingsForm_Cancel","取消"),
				Size=new System.Drawing.Size(100,30),
				Location=new System.Drawing.Point(690,10),
				Anchor=AnchorStyles.Bottom|AnchorStyles.Right
			};
			_btnCancel.Click+=BtnCancel_Click;
			buttonPanel.Controls.Add(_btnCancel);

			// 添加提示标签
			var label = new Label
			{
				Text = ResourceManager.GetString("SettingsForm_ConfigPath", _configFilePath, "配置文件路径: {0}"),
				Dock = DockStyle.Top,
				Height = 30,
				Padding = new Padding(10, 5, 0, 0)
			};
			this.Controls.Add(label);
			this.Controls.SetChildIndex(label,0);
		}

		private void LoadConfigFile()
		{
			try
			{
				if(File.Exists(_configFilePath))
				{
					// 读取文件内容
					string content = File.ReadAllText(_configFilePath);
					// 规范化换行符：统一转换为 CRLF（\r\n），确保 TextBox 正确显示换行
					content=NormalizeLineEndings(content);
					_configTextBox.Text=content;
				} else
				{
					// 如果文件不存在，显示默认配置
					var defaultConfig = new FormattingConfig();
					defaultConfig.Save();
					string content = File.ReadAllText(_configFilePath);
					// 规范化换行符
					content=NormalizeLineEndings(content);
					_configTextBox.Text=content;
				}
			} catch(Exception ex)
			{
				MessageBox.Show(
					ResourceManager.GetString("SettingsForm_LoadError",ex.Message,"加载配置文件失败: {0}"),
					ResourceManager.GetString("SettingsForm_Error","错误"),
					MessageBoxButtons.OK,MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// 规范化换行符，统一转换为 CRLF（\r\n） 处理 LF、CRLF、CR 等不同格式的换行符
		/// </summary>
		private static string NormalizeLineEndings(string text)
		{
			if(string.IsNullOrEmpty(text))
				return text;

			// 先统一处理：将单独的 CR 或 LF 转换为 CRLF
			// 1. 将 CRLF 保持不变（已经是正确的格式）
			// 2. 将单独的 LF (\n) 转换为 CRLF (\r\n)
			// 3. 将单独的 CR (\r) 转换为 CRLF (\r\n)
			// 4. 处理混合情况

			// 方法：先替换所有 CRLF 为临时标记，然后处理单独的 CR 和 LF，最后恢复
			text=text.Replace("\r\n","\n"); // 临时统一为 LF
			text=text.Replace("\r","\n");    // 将单独的 CR 也转换为 LF
			text=text.Replace("\n","\r\n");  // 将所有 LF 转换为 CRLF

			return text;
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
				// 重新加载快捷键
				KeyboardShortcutHelper.ReloadShortcuts();
				Toast.Show(ResourceManager.GetString("Settings_ConfigReloaded","配置已重新加载"),Toast.ToastType.Success);
			}
		}

		private void BtnSave_Click(object sender,EventArgs e)
		{
			try
			{
				// 验证 XML 格式
				var xmlDoc = new System.Xml.XmlDocument();
				xmlDoc.LoadXml(_configTextBox.Text);

				// 规范化换行符，确保保存时使用 CRLF 格式（符合 Windows 和 .editorconfig 要求）
				string contentToSave = NormalizeLineEndings(_configTextBox.Text);

				// 保存到文件
				File.WriteAllText(_configFilePath,contentToSave,System.Text.Encoding.UTF8);

				// 重新加载配置
				FormattingConfig.Reload();
				// 重新加载快捷键
				KeyboardShortcutHelper.ReloadShortcuts();

				Toast.Show(ResourceManager.GetString("Settings_ConfigSaved","配置已保存"),Toast.ToastType.Success);
				this.DialogResult=DialogResult.OK;
				this.Close();
			} catch(System.Xml.XmlException ex)
			{
				MessageBox.Show(
					ResourceManager.GetString("SettingsForm_XMLError",ex.Message,"XML 格式错误: {0}\n\n请检查 XML 语法是否正确。"),
					ResourceManager.GetString("SettingsForm_SaveFailed","保存失败"),
					MessageBoxButtons.OK,MessageBoxIcon.Error);
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

using PPA.Properties;
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Diagnostics;

namespace ToastAPI
{
	/// <summary>
	/// Toast 通知管理器 - 单消息框模式
	/// </summary>
	public static class Toast
	{
		#region Private Fields

		private static readonly object _lock = new();
		private static ToastForm _currentToast;
		private static DateTime _lastShowTime = DateTime.MinValue;

		#endregion Private Fields

		#region Public Enums

		public enum ToastType
		{
			Info,    // 信息
			Success, // 成功
			Warning, // 警告
			Error    // 错误
		}

		#endregion Public Enums

		#region Public Methods

		/// <summary>
		/// 关闭当前显示的通知
		/// </summary>
		public static void CloseCurrentToast()
		{
			lock(_lock)
			{
				Debug.WriteLine($"[Toast通知] 关闭当前通知被调用 | hasCurrent={_currentToast!=null} disposed={_currentToast?.IsDisposed}");

				// 修改：加强空引用检查
				if(_currentToast!=null&&!_currentToast.IsDisposed)
				{
					try
					{
						Debug.WriteLine("[Toast] Forcing close current toast");
						// 修改：同步关闭，避免异步导致的竞态条件
						_currentToast.ForceClose();
						_currentToast=null;
					} catch(Exception ex)
					{
						Debug.WriteLine($"[Toast] Error closing toast: {ex.Message}");
						_currentToast=null;
					}
				}
			}
		}

		/// <summary>
		/// 显示Toast通知（单消息框模式）
		/// </summary>
		/// <param name="message">消息内容</param>
		/// <param name="type">   通知类型</param>
		public static void Show(string message,ToastType type = ToastType.Info,int duration = 99)
		{
			Debug.WriteLine($"[Toast] Show | type={type} durationArg={duration} msgLen={message?.Length}");

			// 修改：防止快速连续调用
			var timeSinceLastShow = DateTime.Now - _lastShowTime;
			if(timeSinceLastShow.TotalMilliseconds<200) // 增加到200ms
			{
				Debug.WriteLine("[Toast通知] 忽略快速重复调用");
				return;
			}
			_lastShowTime=DateTime.Now;

			if(duration==99)
			{
				switch(type)
				{
					case ToastType.Info:
						duration=800;
						break;

					case ToastType.Success:
						duration=1200;
						break;

					case ToastType.Warning:
						duration=1500;
						break;

					case ToastType.Error:
						duration=1800;
						break;
				}
			}

			// 确保在主线程执行
			Debug.WriteLine($"[Toast] Dispatching to UI thread | openForms={Application.OpenForms?.Count}");
			if(Application.OpenForms.Count>0)
			{
				try
				{
					Application.OpenForms[0].Invoke(new Action(() => ShowInternal(message,type,duration)));
				} catch(ObjectDisposedException)
				{
					Debug.WriteLine("[Toast] Main form disposed, ignoring toast");
				} catch(Exception ex)
				{
					Debug.WriteLine($"[Toast] Error invoking ShowInternal: {ex.Message}");
				}
			} else
			{
				ShowInternal(message,type,duration);
			}
		}

		#endregion Public Methods

		#region Internal Methods

		/// <summary>
		/// 当通知关闭时由ToastForm调用
		/// </summary>
		internal static void CurrentToastClosed()
		{
			lock(_lock)
			{
				Debug.WriteLine("[Toast] CurrentToastClosed callback");
				_currentToast=null;
			}
		}

		#endregion Internal Methods

		#region Private Methods

		private static void ShowInternal(string message,ToastType type,int duration)
		{
			lock(_lock)
			{
				Debug.WriteLine($"[Toast] ShowInternal | type={type} duration={duration}");

				// 修改：检查当前窗体状态
				if(_currentToast!=null)
				{
					if(_currentToast.IsDisposed)
					{
						_currentToast=null;
					} else
					{
						CloseCurrentToast();
						// 修改：等待一小段时间确保关闭完成
						System.Threading.Thread.Sleep(50);
					}
				}

				// 创建新通知
				try
				{
					_currentToast=new ToastForm(message,type,duration);
					Debug.WriteLine("[Toast通知] 已创建新的Toast窗体，正在调用Show()");
					_currentToast.Show();
				} catch(Exception ex)
				{
					Debug.WriteLine($"[Toast通知] 创建通知出错: {ex.Message}");
					_currentToast=null;
				}
			}
		}

		#endregion Private Methods
	}

	/// <summary>
	/// 单个Toast通知窗体 - 顶部居中显示
	/// </summary>
	internal class ToastForm:Form
	{
		#region Private Fields

		private readonly int _duration;
		private readonly System.Windows.Forms.Timer _timer;
		private readonly Toast.ToastType _type;
		private bool _isForceClosing = false; // 修改：添加强制关闭标志

		#endregion Private Fields

		#region Public Constructors

		public ToastForm(string message,Toast.ToastType type,int duration)
		{
			_type=type;
			_duration=duration;

			InitializeForm();

			var contentPanel = CreateContentPanel();
			this.Controls.Add(contentPanel);

			AddContentToPanel(contentPanel,message);

			_timer=new System.Windows.Forms.Timer { Interval=duration };
			_timer.Tick+=Timer_Tick;

			SetFixedPosition();

			this.Load+=(s,e) =>
			{
				_timer.Start();
				StartFadeInAnimation();
			};
		}

		#endregion Public Constructors

		#region Public Methods

		// 修改：添加强制关闭方法
		public void ForceClose()
		{
			if(this.IsDisposed) return;

			_isForceClosing=true;
			_timer?.Stop();

			if(this.InvokeRequired)
			{
				this.Invoke(new Action(() =>
				{
					this.Opacity=0;
					this.Close();
					this.Dispose();
				}));
			} else
			{
				this.Opacity=0;
				this.Close();
				this.Dispose();
			}
		}

		#endregion Public Methods

		#region Protected Methods

		protected override CreateParams CreateParams
		{
			get
			{
				CreateParams cp = base.CreateParams;
				cp.ExStyle|=0x02000000; // WS_EX_COMPOSITED
				return cp;
			}
		}

		protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
		{
			base.OnClosing(e);
			if(e.Cancel||_isForceClosing) return;

			e.Cancel=true;
			_timer?.Stop();

			Timer fadeOutTimer = new() { Interval = 15 };
			fadeOutTimer.Tick+=(s,ev) =>
			{
				if(this.Opacity>0.05)
				{
					this.Opacity-=0.08;
				} else
				{
					fadeOutTimer.Stop();
					fadeOutTimer.Dispose();

					e.Cancel=false;
					base.OnClosing(new System.ComponentModel.CancelEventArgs(false));
					this.Dispose();
				}
			};
			fadeOutTimer.Start();
		}

		protected override void OnFormClosed(FormClosedEventArgs e)
		{
			base.OnFormClosed(e);
			_timer?.Stop();
			_timer?.Dispose();

			// 修改：确保通知管理器更新
			try
			{
				Toast.CurrentToastClosed();
			} catch(Exception ex)
			{
				Debug.WriteLine($"[Toast] Error notifying manager: {ex.Message}");
			}
		}

		// 移除鼠标悬停相关函数，因为提示窗显示时间极短

		protected override void OnPaint(PaintEventArgs e)
		{
			if(_isForceClosing) return;

			e.Graphics.SmoothingMode=SmoothingMode.AntiAlias;
			e.Graphics.CompositingQuality=CompositingQuality.HighQuality;
			e.Graphics.InterpolationMode=InterpolationMode.HighQualityBicubic;

			using var path = new GraphicsPath();
			const int radius = 15;
			Rectangle bounds = new(0, 0, this.Width - 1, this.Height - 1);

			path.AddArc(bounds.X,bounds.Y,radius,radius,180,90);
			path.AddArc(bounds.Right-radius,bounds.Y,radius,radius,270,90);
			path.AddArc(bounds.Right-radius,bounds.Bottom-radius,radius,radius,0,90);
			path.AddArc(bounds.X,bounds.Bottom-radius,radius,radius,90,90);
			path.CloseFigure();

			this.Region=new Region(path);

			using var brush = new SolidBrush(this.BackColor);
			e.Graphics.FillPath(brush,path);

			using var pen = new Pen(Color.FromArgb(150, Color.White), 1.5f);
			e.Graphics.DrawPath(pen,path);
		}

		protected override void OnPaintBackground(PaintEventArgs e)
		{
			// 不调用基类方法，避免闪烁
		}

		protected override void WndProc(ref Message m)
		{
			if(m.Msg==0x0002) // WM_DESTROY
			{
				Debug.WriteLine("[Toast通知] 收到WM_DESTROY消息");
			}
			base.WndProc(ref m);
		}

		// 修改：重写Dispose确保资源释放
		protected override void Dispose(bool disposing)
		{
			if(disposing)
			{
				_timer?.Stop();
				_timer?.Dispose();
			}
			base.Dispose(disposing);
		}

		#endregion Protected Methods

		#region Private Methods

		private void AddContentToPanel(Panel panel,string message)
		{
			var iconBox = new PictureBox
			{
				Size = new Size(32, 32),
				Location = new Point(10, (panel.Height - 32) / 2),
				SizeMode = PictureBoxSizeMode.Zoom,
				BackColor = Color.Transparent,
				Image = GetIconForType(),
			};

			panel.Controls.Add(iconBox);

			var label = new Label
			{
				Text = message,
				AutoSize = false,
				Size = new Size(panel.Width - 60, panel.Height - 20),
				Location = new Point(50, 10),
				TextAlign = ContentAlignment.MiddleLeft,
				ForeColor = Color.White,
				BackColor = Color.Transparent,
				Font = new Font("Segoe UI", 10, FontStyle.Regular),
				UseCompatibleTextRendering = false
			};

			panel.Controls.Add(label);
		}

		private Panel CreateContentPanel()
		{
			var panel = new DoubleBufferedPanel
			{
				Dock = DockStyle.Fill,
				Padding = new Padding(10),
				BackColor = Color.Transparent
			};
			return panel;
		}

		private Image GetIconForType()
		{
			return _type switch
			{
				Toast.ToastType.Success => Resources.Infom,
				Toast.ToastType.Warning => Resources.warn,
				Toast.ToastType.Error => Resources.erro,
				_ => Resources.Infom,
			};
		}

		private Color GetSolidColorForType()
		{
			bool isDarkTheme = System.Windows.Forms.Control.DefaultBackColor.GetBrightness() < 0.5;

			return _type switch
			{
				Toast.ToastType.Success => isDarkTheme ?
										Color.FromArgb(76,175,80) :
										Color.FromArgb(56,142,60),
				Toast.ToastType.Warning => isDarkTheme ?
										Color.FromArgb(255,193,7) :
										Color.FromArgb(230,174,0),
				Toast.ToastType.Error => isDarkTheme ?
										Color.FromArgb(244,67,54) :
										Color.FromArgb(220,60,48),
				_ => isDarkTheme ?
										Color.FromArgb(33,150,243) :
										Color.FromArgb(28,130,223),
			};
		}

		private void InitializeForm()
		{
			this.FormBorderStyle=FormBorderStyle.None;
			this.ShowInTaskbar=false;
			this.TopMost=true;
			this.StartPosition=FormStartPosition.Manual;
			this.Size=new Size(350,70);
			this.Padding=new Padding(0);
			this.Font=new Font("Segoe UI",10);
			this.DoubleBuffered=true;
			this.Opacity=0;
			this.BackColor=GetSolidColorForType();
		}

		private void SetFixedPosition()
		{
			Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
			int leftPosition = workingArea.Left + ((workingArea.Width - this.Width) / 2);
			int topPosition = workingArea.Top + 20;
			this.Location=new Point(leftPosition,topPosition);
		}

		private void StartFadeInAnimation()
		{
			Timer fadeInTimer = new() { Interval = 30 };

			fadeInTimer.Tick+=(s,e) =>
			{
				if(_isForceClosing)
				{
					fadeInTimer.Stop();
					fadeInTimer.Dispose();
					return;
				}

				if(this.Opacity<0.9)
				{
					this.Opacity+=0.05;
				} else
				{
					this.Opacity=0.9;
					fadeInTimer.Stop();
					fadeInTimer.Dispose();
				}
			};
			fadeInTimer.Start();
		}

		private void Timer_Tick(object sender,EventArgs e)
		{
			// 直接关闭提示窗，不再检查鼠标悬停状态
			_timer.Stop();
			this.Close();
		}

		#endregion Private Methods
	}

	internal class DoubleBufferedPanel:Panel
	{
		public DoubleBufferedPanel()
		{
			this.DoubleBuffered=true;
		}
	}
}

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.IO;

using RAD.ClipMon.Win32;
using RAD.Windows;

namespace RAD.ClipMon
{
	/// <summary>
    /// Code Page 1252 Fixer
    /// Copyright (c) Sam Pospischil 2011
    /// <br/>
    /// Based on Clipboard Monitor Example 
	/// Copyright (c) Ross Donald 2003
	/// ross@radsoftware.com.au
	/// http://www.radsoftware.com.au
	/// <br/>
	/// Monitors the clipboard for changes, and replaces characters from Windows CP 1252
    /// with those from the low ASCII range. Configurable via tray menu as follows:
    /// (also a :TODO: list)
    /// - Toggle replace 'smart quotes' (accented single and double quotes, long dashes and ellipsis)
    /// - Toggle replace smart quotes as HTML entites, preserving them rather than converting down
    /// - Toggle replace any characters which have named HTML entities
    /// - Toggle replace all high-ASCII characters with entity numbers
    /// - Toggle automatic conversion of RTF text to plaintext
	/// </summary>
	/// <remarks>
	/// </remarks>
	public class frmMain : System.Windows.Forms.Form
	{
        #region Constants



		#endregion


		#region Fields

		private System.ComponentModel.IContainer components;

        private System.Windows.Forms.MainMenu menuMain;
        private System.Windows.Forms.RichTextBox ctlClipboardText;
		protected System.Windows.Forms.ContextMenu cmnuTray;
		private System.Windows.Forms.MenuItem itmExit;
		private System.Windows.Forms.MenuItem itmHide;
		private System.Windows.Forms.MenuItem itmSep2;
        private System.Windows.Forms.MenuItem itmSep1;
		private RAD.Windows.NotificationAreaIcon notAreaIcon;

		#endregion


		#region Constructors

		public frmMain()
		{
			InitializeComponent();	
			notAreaIcon.Visible = true;
		}

		#endregion


		#region Properties - Public



		#endregion

        #region Properties - Private

        private String _rawClip;
        private String _currentClip;
        private String _modifiedClip;

        private Queue _history = new Queue();
        private IntPtr _ClipboardViewerNext;

        #endregion


        #region Methods - Private

        /// <summary>
		/// Register this form as a Clipboard Viewer application
		/// </summary>
		private void RegisterClipboardViewer()
		{
			_ClipboardViewerNext = Win32.User32.SetClipboardViewer(this.Handle);
		}

		/// <summary>
		/// Remove this form from the Clipboard Viewer list
		/// </summary>
		private void UnregisterClipboardViewer()
		{
			Win32.User32.ChangeClipboardChain(this.Handle, _ClipboardViewerNext);
		}

		/// <summary>
		/// Search the clipboard contents for configured characters and translate them
		/// </summary>
		/// <param name="iData">The current clipboard contents</param>
		/// <returns>true if characters were replaced, false otherwise</returns>
        private bool replaceQuotes(String clipContents)
        {
            bool found = false;

            // :TODO: pretty much everything
            _modifiedClip = _currentClip;

            return found;
        }

        private String getClipAsString(IDataObject iData)
		{
            String contents = "";
			//
			// If it is not text then quit
			// cannot search bitmap etc
			//
            try 
			{
			    if (iData.GetDataPresent(DataFormats.Rtf))
			    {
                    _rawClip = (string)iData.GetData(DataFormats.Rtf);

                    // convert RTF data to plainText using a RichTextBox
                    System.Windows.Forms.RichTextBox rtBox = new System.Windows.Forms.RichTextBox();
                    rtBox.Rtf = (string)_rawClip;
                    contents = rtBox.Text;

				    setNotificationTooltip("RTF copied");
			    }
                else if (iData.GetDataPresent(DataFormats.Text))
                {
                    _rawClip = (string)iData.GetData(DataFormats.Text);
                    contents = _rawClip;
                    setNotificationTooltip("Text copied");
                }
            }
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
            return contents;
		}

		/// <summary>
		/// Build the system tray menu from the hyperlink list
		/// </summary>
		private void ContextMenuBuild()
		{
			cmnuTray.MenuItems.Clear();

			// :TODO: add mode options

			cmnuTray.MenuItems.Add("-");
			cmnuTray.MenuItems.Add("Cancel Menu", new EventHandler(itmCancelMenu_Click));
			cmnuTray.MenuItems.Add("-");
			cmnuTray.MenuItems.Add(itmHide.Text, new EventHandler(itmHide_Click));
			cmnuTray.MenuItems.Add("-");
			cmnuTray.MenuItems.Add("E&xit", new EventHandler(itmExit_Click));
		}

        private void setNotificationTooltip(String tt)
        {
            notAreaIcon.Tooltip = tt;
        }

		/// <summary>
		/// Show the clipboard contents in the window 
		/// and show the notification balloon if a link is found
		/// </summary>
		private void ProcessClipboard()
		{
			//
			// Data on the clipboard uses the 
			// IDataObject interface
			//
			IDataObject iData = new DataObject();

			try
			{
				iData = Clipboard.GetDataObject();
			}
			catch (System.Runtime.InteropServices.ExternalException externEx)
			{
				// Copying a field definition in Access 2002 causes this sometimes?
				Debug.WriteLine("InteropServices.ExternalException: {0}", externEx.Message);
				return;
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.ToString());
				return;
			}

            _currentClip = getClipAsString(iData);
            
            // show current clipboard text even if there was no match
            if (_currentClip != "")
            {
                if (iData.GetDataPresent(DataFormats.Rtf))
                {
                    ctlClipboardText.Rtf = _rawClip;   // but show the original RTF formatted string where appropriate
                }
                else
                {
                    ctlClipboardText.Text = _rawClip;
                }
            }

            if (replaceQuotes(_currentClip))
			{
				// bad quotes were found and have be purged, so update the text with the fixed one
                ctlClipboardText.Text = _modifiedClip;

				notAreaIcon.BalloonDisplay(NotificationAreaIcon.NOTIFYICONdwInfoFlags.NIIF_INFO, "Characters fixed", "Click to view or modify the clipboard contents");
			}
		}

		#endregion


		#region Methods - Public

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new frmMain());
		}


		protected override void WndProc(ref Message m)
		{
			switch ((Win32.Msgs)m.Msg)
			{
					//
					// The WM_DRAWCLIPBOARD message is sent to the first window 
					// in the clipboard viewer chain when the content of the 
					// clipboard changes. This enables a clipboard viewer 
					// window to display the new content of the clipboard. 
					//
				case Win32.Msgs.WM_DRAWCLIPBOARD:
					
					Debug.WriteLine("WindowProc DRAWCLIPBOARD: " + m.Msg, "WndProc");

                    ProcessClipboard();

					//
					// Each window that receives the WM_DRAWCLIPBOARD message 
					// must call the SendMessage function to pass the message 
					// on to the next window in the clipboard viewer chain.
					//
					Win32.User32.SendMessage(_ClipboardViewerNext, m.Msg, m.WParam, m.LParam);
					break;


					//
					// The WM_CHANGECBCHAIN message is sent to the first window 
					// in the clipboard viewer chain when a window is being 
					// removed from the chain. 
					//
				case Win32.Msgs.WM_CHANGECBCHAIN:
					Debug.WriteLine("WM_CHANGECBCHAIN: lParam: " + m.LParam, "WndProc");

					// When a clipboard viewer window receives the WM_CHANGECBCHAIN message, 
					// it should call the SendMessage function to pass the message to the 
					// next window in the chain, unless the next window is the window 
					// being removed. In this case, the clipboard viewer should save 
					// the handle specified by the lParam parameter as the next window in the chain. 

					//
					// wParam is the Handle to the window being removed from 
					// the clipboard viewer chain 
					// lParam is the Handle to the next window in the chain 
					// following the window being removed. 
					if (m.WParam == _ClipboardViewerNext)
					{
						//
						// If wParam is the next clipboard viewer then it
						// is being removed so update pointer to the next
						// window in the clipboard chain
						//
						_ClipboardViewerNext = m.LParam;
					}
					else
					{
						Win32.User32.SendMessage(_ClipboardViewerNext, m.Msg, m.WParam, m.LParam);
					}
					break;

				default:
					//
					// Let the form process the messages that we are
					// not interested in
					//
					base.WndProc(ref m);
					break;

			}

		}

		#endregion


		#region Event Handlers - Menu

		private void itmExit_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void itmHide_Click(object sender, System.EventArgs e)
		{
			this.Visible = (! this.Visible);
			itmHide.Text = this.Visible ? "Hide" : "Show";

			if (this.Visible == true)
			{
				if (this.WindowState == FormWindowState.Minimized)
				{
					this.WindowState = FormWindowState.Normal;
				}
			}
		}

		private void itmCancelMenu_Click(object sender, System.EventArgs e)
		{
			// Do nothing - Cancel the menu
		}

		private void frmMain_Resize(object sender, System.EventArgs e)
		{
			if ((this.WindowState == FormWindowState.Minimized) && (this.Visible == true))
			{
				// hide when minimised
				this.Visible = false;
				itmHide.Text = "Show";
			}
		}

		#endregion


		#region Event Handlers - Internal

		private void frmMain_Load(object sender, System.EventArgs e)
		{
			RegisterClipboardViewer();
		}

		private void frmMain_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			UnregisterClipboardViewer();
		}

		private void notAreaIcon_BalloonClick(object sender, System.EventArgs e)
		{
			notAreaIcon.ContextMenuDisplay();
		}

		#endregion


		#region IDisposable Implementation

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#endregion


		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.menuMain = new System.Windows.Forms.MainMenu(this.components);
            this.notAreaIcon = new RAD.Windows.NotificationAreaIcon(this.components);
            this.cmnuTray = new System.Windows.Forms.ContextMenu();
            this.itmSep1 = new System.Windows.Forms.MenuItem();
            this.itmHide = new System.Windows.Forms.MenuItem();
            this.itmSep2 = new System.Windows.Forms.MenuItem();
            this.itmExit = new System.Windows.Forms.MenuItem();
            this.ctlClipboardText = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // notAreaIcon
            // 
            this.notAreaIcon.ContextMenu = this.cmnuTray;
            this.notAreaIcon.DisplayMenuOnLeftClick = true;
            this.notAreaIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notAreaIcon.Icon")));
            this.notAreaIcon.Tooltip = "CP-1252 Fixer";
            this.notAreaIcon.Visible = false;
            this.notAreaIcon.BalloonClick += new System.EventHandler(this.notAreaIcon_BalloonClick);
            // 
            // cmnuTray
            // 
            this.cmnuTray.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.itmSep1,
            this.itmHide,
            this.itmSep2,
            this.itmExit});
            // 
            // itmSep1
            // 
            this.itmSep1.Index = 0;
            this.itmSep1.Text = "-";
            // 
            // itmHide
            // 
            this.itmHide.Index = 1;
            this.itmHide.Text = "Hide";
            this.itmHide.Click += new System.EventHandler(this.itmHide_Click);
            // 
            // itmSep2
            // 
            this.itmSep2.Index = 2;
            this.itmSep2.Text = "-";
            // 
            // itmExit
            // 
            this.itmExit.Index = 3;
            this.itmExit.MergeOrder = 1000;
            this.itmExit.Text = "E&xit";
            // 
            // ctlClipboardText
            // 
            this.ctlClipboardText.DetectUrls = false;
            this.ctlClipboardText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ctlClipboardText.Location = new System.Drawing.Point(0, 0);
            this.ctlClipboardText.Name = "ctlClipboardText";
            this.ctlClipboardText.ReadOnly = true;
            this.ctlClipboardText.Size = new System.Drawing.Size(348, 273);
            this.ctlClipboardText.TabIndex = 0;
            this.ctlClipboardText.Text = "";
            this.ctlClipboardText.WordWrap = false;
            // 
            // frmMain
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
            this.ClientSize = new System.Drawing.Size(348, 273);
            this.Controls.Add(this.ctlClipboardText);
            this.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(100, 100);
            this.Menu = this.menuMain;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "CP-1252 Fixer";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmMain_Closing);
            this.Resize += new System.EventHandler(this.frmMain_Resize);
            this.ResumeLayout(false);

		}
		#endregion

		

	}
}

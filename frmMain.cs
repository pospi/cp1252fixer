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
	/// Clipboard Monitor Example 
	/// Copyright (c) Ross Donald 2003
	/// ross@radsoftware.com.au
	/// http://www.radsoftware.com.au
	/// <br/>
	/// Demonstrates how to create a clipboard monitor in C#. Whenever an item is copied
	/// to the clipboard by any application this form will be notified by a call to 
	/// the WindowProc method with the WM_DRAWCLIPBOARD message allowing this form to
	/// read the contents of the clipboard and perform some processing.
	/// </summary>
	/// <remarks>
	/// This application has some functionality beyond a simple example. When an item is copied
	/// to the clipboard this application looks for hyperlinks, unc paths or email addresses 
	/// then displays a balloon dialog (Windows XP only) showing the link that was found.
	/// The icon in the system tray area can be clicked to display a menu of the found links.
	/// <br/>
	/// This source code is a work in progress and comes without warranty expressed or implied.
	/// It is an attempt to demonstrate a concept, not to be a finished application.
	/// </remarks>
	public class frmMain : System.Windows.Forms.Form
	{
		#region Clipboard Formats

		string[] formatsAll = new string[] 
		{
			DataFormats.Bitmap,
			DataFormats.CommaSeparatedValue,
			DataFormats.Dib,
			DataFormats.Dif,
			DataFormats.EnhancedMetafile,
			DataFormats.FileDrop,
			DataFormats.Html,
			DataFormats.Locale,
			DataFormats.MetafilePict,
			DataFormats.OemText,
			DataFormats.Palette,
			DataFormats.PenData,
			DataFormats.Riff,
			DataFormats.Rtf,
			DataFormats.Serializable,
			DataFormats.StringFormat,
			DataFormats.SymbolicLink,
			DataFormats.Text,
			DataFormats.Tiff,
			DataFormats.UnicodeText,
			DataFormats.WaveAudio
		};

		string[] formatsAllDesc = new String[] 
		{
			"Bitmap",
			"CommaSeparatedValue",
			"Dib",
			"Dif",
			"EnhancedMetafile",
			"FileDrop",
			"Html",
			"Locale",
			"MetafilePict",
			"OemText",
			"Palette",
			"PenData",
			"Riff",
			"Rtf",
			"Serializable",
			"StringFormat",
			"SymbolicLink",
			"Text",
			"Tiff",
			"UnicodeText",
			"WaveAudio"
		};

		#endregion


		#region Constants



		#endregion


		#region Fields

		private System.ComponentModel.IContainer components;

		private System.Windows.Forms.MainMenu menuMain;
		private System.Windows.Forms.MenuItem mnuFormats;
		private System.Windows.Forms.RichTextBox ctlClipboardText;
		private System.Windows.Forms.MenuItem mnuSupported;
		protected System.Windows.Forms.ContextMenu cmnuTray;
		private System.Windows.Forms.MenuItem itmExit;
		private System.Windows.Forms.MenuItem itmHide;
		private System.Windows.Forms.MenuItem itmSep2;
		private System.Windows.Forms.MenuItem itmSep1;
		private System.Windows.Forms.MenuItem itmHyperlink;
		private System.Windows.Forms.MenuItem itmSystray;
		private RAD.Windows.NotificationAreaIcon notAreaIcon;

		IntPtr _ClipboardViewerNext;
		Queue _hyperlink = new Queue(); 		

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
		/// Build a menu listing the formats supported by the contents of the clipboard
		/// </summary>
		/// <param name="iData">The current clipboard data object</param>
		private void FormatMenuBuild(IDataObject iData)
		{
			string[] astrFormatsNative = iData.GetFormats(false);
			string[] astrFormatsAll = iData.GetFormats(true);

			Hashtable formatList = new Hashtable(10);

			mnuFormats.MenuItems.Clear();

			for (int i = 0; i <= astrFormatsAll.GetUpperBound(0); i++)
			{
				formatList.Add(astrFormatsAll[i], "Converted");
			}

			for (int i = 0; i <= astrFormatsNative.GetUpperBound(0); i++)
			{
				if (formatList.Contains(astrFormatsNative[i]))
				{
					formatList[astrFormatsNative[i]] = "Native/Converted";
				}
				else
				{
					formatList.Add(astrFormatsNative[i], "Native");
				}
			}

			foreach (DictionaryEntry item in formatList) 
			{
				MenuItem itmNew = new MenuItem(item.Key.ToString() + "\t" + item.Value.ToString());
				mnuFormats.MenuItems.Add(itmNew);
			}
		}

		/// <summary>
		/// list the formats that are supported from the default clipboard formats.
		/// </summary>
		/// <param name="iData">The current clipboard contents</param>
		private void SupportedMenuBuild(IDataObject iData)
		{
			mnuSupported.MenuItems.Clear();
		
			for (int i = 0; i <= formatsAll.GetUpperBound(0); i++)
			{
				MenuItem itmFormat = new MenuItem(formatsAllDesc[i]);
				//
				// Get supported formats
				//
				if (iData.GetDataPresent(formatsAll[i]))
				{
					itmFormat.Checked = true;
				}
				mnuSupported.MenuItems.Add(itmFormat);
		
			}
		}

		/// <summary>
		/// Search the clipboard contents for hyperlinks and unc paths, etc
		/// </summary>
		/// <param name="iData">The current clipboard contents</param>
		/// <returns>true if new links were found, false otherwise</returns>
		private bool ClipboardSearch(IDataObject iData)
		{
			bool FoundNewLinks = false;
			//
			// If it is not text then quit
			// cannot search bitmap etc
			//
			if (!iData.GetDataPresent(DataFormats.Text))
			{
				return false; 
			}

			string strClipboardText;

			try 
			{
				 strClipboardText = (string)iData.GetData(DataFormats.Text);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
				return false;
			}
			
			// Hyperlinks e.g. http://www.server.com/folder/file.aspx
			Regex rxURL = new Regex(@"(\b(?:http|https|ftp|file)://[^\s]+)", RegexOptions.IgnoreCase);
			rxURL.Match(strClipboardText);

			foreach (Match rm in rxURL.Matches(strClipboardText))
			{
				if(!_hyperlink.Contains(rm.ToString()))
				{
					_hyperlink.Enqueue(rm.ToString());
					FoundNewLinks = true;
				}
			}

			// Files and folders - \\servername\foldername\
			// TODO needs work
			Regex rxFile = new Regex(@"(\b\w:\\[^ ]*)", RegexOptions.IgnoreCase);
			rxFile.Match(strClipboardText);

			foreach (Match rm in rxFile.Matches(strClipboardText))
			{
				if(!_hyperlink.Contains(rm.ToString()))
				{
					_hyperlink.Enqueue(rm.ToString());
					FoundNewLinks = true;
				}
			}

			// UNC Files 
			// TODO needs work
			Regex rxUNC = new Regex(@"(\\\\[^\s/:\*\?\" + "\"" + @"\<\>\|]+)", RegexOptions.IgnoreCase);
			rxUNC.Match(strClipboardText);

			foreach (Match rm in rxUNC.Matches(strClipboardText))
			{
				if(!_hyperlink.Contains(rm.ToString()))
				{
					_hyperlink.Enqueue(rm.ToString());
					FoundNewLinks = true;
				}
			}

			// UNC folders
			// TODO needs work
			Regex rxUNCFolder = new Regex(@"(\\\\[^\s/:\*\?\" + "\"" + @"\<\>\|]+\\)", RegexOptions.IgnoreCase);
			rxUNCFolder.Match(strClipboardText);

			foreach (Match rm in rxUNCFolder.Matches(strClipboardText))
			{
				if(!_hyperlink.Contains(rm.ToString()))
				{
					_hyperlink.Enqueue(rm.ToString());
					FoundNewLinks = true;
				}
			}

			// Email Addresses
			Regex rxEmailAddress = new Regex(@"([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)", RegexOptions.IgnoreCase);
			rxEmailAddress.Match(strClipboardText);

			foreach (Match rm in rxEmailAddress.Matches(strClipboardText))
			{
				if(!_hyperlink.Contains(rm.ToString()))
				{
					_hyperlink.Enqueue("mailto:" + rm.ToString());
					FoundNewLinks = true;
				}
			}

			return FoundNewLinks;
		}

		/// <summary>
		/// Build the system tray menu from the hyperlink list
		/// </summary>
		private void ContextMenuBuild()
		{
			//
			// Only show the last 10 items
			//
			while (_hyperlink.Count > 10)
			{
				_hyperlink.Dequeue();
			}

			cmnuTray.MenuItems.Clear();

			foreach (string objLink in _hyperlink)
			{
				cmnuTray.MenuItems.Add(objLink.ToString(), new EventHandler(itmHyperlink_Click));
			}
			cmnuTray.MenuItems.Add("-");
			cmnuTray.MenuItems.Add("Cancel Menu", new EventHandler(itmCancelMenu_Click));
			cmnuTray.MenuItems.Add("-");
			cmnuTray.MenuItems.Add(itmHide.Text, new EventHandler(itmHide_Click));
			cmnuTray.MenuItems.Add("-");
			cmnuTray.MenuItems.Add("E&xit", new EventHandler(itmExit_Click));
		}


		/// <summary>
		/// Called when an item is chosen from the menu
		/// </summary>
		/// <param name="pstrLink">The link that was clicked</param>
		private void OpenLink(string pstrLink)
		{
			try
			{
				//
				// Run the link
				//

				// TODO needs more work to check for missing files etc
				System.Diagnostics.Process.Start(pstrLink);
			}
			catch (Exception ex)
			{
				MessageBox.Show(this, ex.ToString(), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}

		}



		/// <summary>
		/// Show the clipboard contents in the window 
		/// and show the notification balloon if a link is found
		/// </summary>
		private void GetClipboardData()
		{
			//
			// Data on the clipboard uses the 
			// IDataObject interface
			//
			IDataObject iData = new DataObject();  
			string strText = "clipmon";

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
					
			// 
			// Get RTF if it is present
			//
			if (iData.GetDataPresent(DataFormats.Rtf))
			{
				ctlClipboardText.Rtf = (string)iData.GetData(DataFormats.Rtf);
						
				if(iData.GetDataPresent(DataFormats.Text))
				{
					strText = "RTF";
				}
			}
			else
			{
				// 
				// Get Text if it is present
				//
				if(iData.GetDataPresent(DataFormats.Text))
				{
					ctlClipboardText.Text = (string)iData.GetData(DataFormats.Text);
							
					strText = "Text"; 

					Debug.WriteLine((string)iData.GetData(DataFormats.Text));
				}
				else
				{
					//
					// Only show RTF or TEXT
					//
					ctlClipboardText.Text = "(cannot display this format)";
				}
			}

			notAreaIcon.Tooltip = strText;

			if( ClipboardSearch(iData) )
			{
				//
				// Found some new links
				//
				System.Text.StringBuilder strBalloon = new System.Text.StringBuilder(100);

				foreach (string objLink in _hyperlink)
				{
					strBalloon.Append(objLink.ToString()  + "\n");
				}

				FormatMenuBuild(iData);
				SupportedMenuBuild(iData);					
				ContextMenuBuild();

				if (_hyperlink.Count > 0)
				{
					notAreaIcon.BalloonDisplay(NotificationAreaIcon.NOTIFYICONdwInfoFlags.NIIF_INFO, "Links", strBalloon.ToString());
				}
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

					GetClipboardData();

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

		private void itmHyperlink_Click(object sender, System.EventArgs e)
		{
			MenuItem itmHL = (MenuItem)sender;

			OpenLink(itmHL.Text);
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
			if(_hyperlink.Count == 1)
			{
				string strItem = (string)_hyperlink.ToArray()[0];

				// Only one link so open it
				OpenLink(strItem);
			}
			else
			{
				notAreaIcon.ContextMenuDisplay();
			}
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmMain));
			this.menuMain = new System.Windows.Forms.MainMenu();
			this.mnuFormats = new System.Windows.Forms.MenuItem();
			this.mnuSupported = new System.Windows.Forms.MenuItem();
			this.notAreaIcon = new RAD.Windows.NotificationAreaIcon(this.components);
			this.cmnuTray = new System.Windows.Forms.ContextMenu();
			this.itmSystray = new System.Windows.Forms.MenuItem();
			this.itmHyperlink = new System.Windows.Forms.MenuItem();
			this.itmSep1 = new System.Windows.Forms.MenuItem();
			this.itmHide = new System.Windows.Forms.MenuItem();
			this.itmSep2 = new System.Windows.Forms.MenuItem();
			this.itmExit = new System.Windows.Forms.MenuItem();
			this.ctlClipboardText = new System.Windows.Forms.RichTextBox();
			this.SuspendLayout();
			// 
			// menuMain
			// 
			this.menuMain.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					 this.mnuFormats,
																					 this.mnuSupported});
			// 
			// mnuFormats
			// 
			this.mnuFormats.Index = 0;
			this.mnuFormats.Text = "Formats";
			// 
			// mnuSupported
			// 
			this.mnuSupported.Index = 1;
			this.mnuSupported.Text = "Supported";
			// 
			// notAreaIcon
			// 
			this.notAreaIcon.ContextMenu = this.cmnuTray;
			this.notAreaIcon.DisplayMenuOnLeftClick = true;
			this.notAreaIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notAreaIcon.Icon")));
			this.notAreaIcon.Tooltip = "Clip Monitor";
			this.notAreaIcon.Visible = false;
			this.notAreaIcon.BalloonClick += new System.EventHandler(this.notAreaIcon_BalloonClick);
			// 
			// cmnuTray
			// 
			this.cmnuTray.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					 this.itmSystray,
																					 this.itmHyperlink,
																					 this.itmSep1,
																					 this.itmHide,
																					 this.itmSep2,
																					 this.itmExit});
			// 
			// itmSystray
			// 
			this.itmSystray.Index = 0;
			this.itmSystray.Text = "C:\\Temp\\SysTray";
			this.itmSystray.Click += new System.EventHandler(this.itmHyperlink_Click);
			// 
			// itmHyperlink
			// 
			this.itmHyperlink.DefaultItem = true;
			this.itmHyperlink.Index = 1;
			this.itmHyperlink.Text = "http://localhost/footprint/";
			this.itmHyperlink.Click += new System.EventHandler(this.itmHyperlink_Click);
			// 
			// itmSep1
			// 
			this.itmSep1.Index = 2;
			this.itmSep1.Text = "-";
			// 
			// itmHide
			// 
			this.itmHide.Index = 3;
			this.itmHide.Text = "Hide";
			this.itmHide.Click += new System.EventHandler(this.itmHide_Click);
			// 
			// itmSep2
			// 
			this.itmSep2.Index = 4;
			this.itmSep2.Text = "-";
			// 
			// itmExit
			// 
			this.itmExit.Index = 5;
			this.itmExit.MergeOrder = 1000;
			this.itmExit.Text = "E&xit";
			// 
			// ctlClipboardText
			// 
			this.ctlClipboardText.DetectUrls = false;
			this.ctlClipboardText.Dock = System.Windows.Forms.DockStyle.Fill;
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
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.ctlClipboardText});
			this.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Location = new System.Drawing.Point(100, 100);
			this.Menu = this.menuMain;
			this.Name = "frmMain";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Clipboard Monitor Sample from www.radsoftware.com.au";
			this.Resize += new System.EventHandler(this.frmMain_Resize);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmMain_Closing);
			this.Load += new System.EventHandler(this.frmMain_Load);
			this.ResumeLayout(false);

		}
		#endregion

		

	}
}

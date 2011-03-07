using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using RAD.ClipMon.Win32;
using RAD.Windows;

namespace pospi.CP1252
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
        #region Character Mappings

        Dictionary<int, String> smartQuotesTranslate = new Dictionary<int, String>()
        {
            { 0x201A, ","},
            { 0x201E, ",,"},
            { 0x2018, "\'"},
            { 0x2019, "\'"},
            { 0x201C, "\""},
            { 0x201D, "\""},
            { 0x2022, "-"},
            { 0x2013, "-"},
            { 0x2014, "-"},
            { 0x2026, "..."}
        };

        Dictionary<int, String> smartQuotesEscape = new Dictionary<int, String>()
        {
            { 0x201A, "&sbquo;"},
            { 0x201E, "&bdquo;"},
            { 0x2018, "&lsquo;"},
            { 0x2019, "&rsquo;"},
            { 0x201C, "&ldquo;"},
            { 0x201D, "&rdquo;"},
            { 0x2022, "&bull;"},
            { 0x2013, "&ndash;"},
            { 0x2014, "&mdash;"},
            { 0x2026, "&hellip;"},
            { 188, "&frac14;"},       // we also escape fractions as they are auto-inserted by Word
            { 189, "&frac12;"},
            { 190, "&frac34;"}
        };

        Dictionary<int, String> namedEntityEscape = new Dictionary<int, String>()
        {
            { 160, "&nbsp;"},
            { 161, "&iexcl;"},
            { 162, "&cent;"},
            { 163, "&pound;"},
            { 164, "&curren;"},
            { 165, "&yen;"},
            { 166, "&brvbar;"},
            { 167, "&sect;"},
            { 168, "&uml;"},
            { 169, "&copy;"},
            { 170, "&ordf;"},
            { 171, "&laquo;"},
            { 172, "&not;"},
            { 173, "&shy;"},
            { 174, "&reg;"},
            { 175, "&macr;"},
            { 176, "&deg;"},
            { 177, "&plusmn;"},
            { 178, "&sup2;"},
            { 179, "&sup3;"},
            { 180, "&acute;"},
            { 181, "&micro;"},
            { 182, "&para;"},
            { 183, "&middot;"},
            { 184, "&cedil;"},
            { 185, "&sup1;"},
            { 186, "&ordm;"},
            { 187, "&raquo;"},
            { 188, "&frac14;"},
            { 189, "&frac12;"},
            { 190, "&frac34;"},
            { 191, "&iquest;"},
            { 192, "&Agrave;"},
            { 193, "&Aacute;"},
            { 194, "&Acirc;"},
            { 195, "&Atilde;"},
            { 196, "&Auml;"},
            { 197, "&Aring;"},
            { 198, "&AElig;"},
            { 199, "&Ccedil;"},
            { 200, "&Egrave;"},
            { 201, "&Eacute;"},
            { 202, "&Ecirc;"},
            { 203, "&Euml;"},
            { 204, "&Igrave;"},
            { 205, "&Iacute;"},
            { 206, "&Icirc;"},
            { 207, "&Iuml;"},
            { 208, "&ETH;"},
            { 209, "&Ntilde;"},
            { 210, "&Ograve;"},
            { 211, "&Oacute;"},
            { 212, "&Ocirc;"},
            { 213, "&Otilde;"},
            { 214, "&Ouml;"},
            { 215, "&times;"},
            { 216, "&Oslash;"},
            { 217, "&Ugrave;"},
            { 218, "&Uacute;"},
            { 219, "&Ucirc;"},
            { 220, "&Uuml;"},
            { 221, "&Yacute;"},
            { 222, "&THORN;"},
            { 223, "&szlig;"},
            { 224, "&agrave;"},
            { 225, "&aacute;"},
            { 226, "&acirc;"},
            { 227, "&atilde;"},
            { 228, "&auml;"},
            { 229, "&aring;"},
            { 230, "&aelig;"},
            { 231, "&ccedil;"},
            { 232, "&egrave;"},
            { 233, "&eacute;"},
            { 234, "&ecirc;"},
            { 235, "&euml;"},
            { 236, "&igrave;"},
            { 237, "&iacute;"},
            { 238, "&icirc;"},
            { 239, "&iuml;"},
            { 240, "&eth;"},
            { 241, "&ntilde;"},
            { 242, "&ograve;"},
            { 243, "&oacute;"},
            { 244, "&ocirc;"},
            { 245, "&otilde;"},
            { 246, "&ouml;"},
            { 247, "&divide;"},
            { 248, "&oslash;"},
            { 249, "&ugrave;"},
            { 250, "&uacute;"},
            { 251, "&ucirc;"},
            { 252, "&uuml;"},
            { 253, "&yacute;"},
            { 254, "&thorn;"},
            { 255, "&yuml;"},
            { 402, "&fnof;"},
            { 913, "&Alpha;"},
            { 914, "&Beta;"},
            { 915, "&Gamma;"},
            { 916, "&Delta;"},
            { 917, "&Epsilon;"},
            { 918, "&Zeta;"},
            { 919, "&Eta;"},
            { 920, "&Theta;"},
            { 921, "&Iota;"},
            { 922, "&Kappa;"},
            { 923, "&Lambda;"},
            { 924, "&Mu;"},
            { 925, "&Nu;"},
            { 926, "&Xi;"},
            { 927, "&Omicron;"},
            { 928, "&Pi;"},
            { 929, "&Rho;"},
            { 931, "&Sigma;"},
            { 932, "&Tau;"},
            { 933, "&Upsilon;"},
            { 934, "&Phi;"},
            { 935, "&Chi;"},
            { 936, "&Psi;"},
            { 937, "&Omega;"},
            { 945, "&alpha;"},
            { 946, "&beta;"},
            { 947, "&gamma;"},
            { 948, "&delta;"},
            { 949, "&epsilon;"},
            { 950, "&zeta;"},
            { 951, "&eta;"},
            { 952, "&theta;"},
            { 953, "&iota;"},
            { 954, "&kappa;"},
            { 955, "&lambda;"},
            { 956, "&mu;"},
            { 957, "&nu;"},
            { 958, "&xi;"},
            { 959, "&omicron;"},
            { 960, "&pi;"},
            { 961, "&rho;"},
            { 962, "&sigmaf;"},
            { 963, "&sigma;"},
            { 964, "&tau;"},
            { 965, "&upsilon;"},
            { 966, "&phi;"},
            { 967, "&chi;"},
            { 968, "&psi;"},
            { 969, "&omega;"},
            { 977, "&thetasym;"},
            { 978, "&upsih;"},
            { 982, "&piv;"},
            //{ 8226, "&bull;"},
            //{ 8230, "&hellip;"},
            { 8242, "&prime;"},
            { 8243, "&Prime;"},
            { 8254, "&oline;"},
            { 8260, "&frasl;"},
            { 8472, "&weierp;"},
            { 8465, "&image;"},
            { 8476, "&real;"},
            { 8482, "&trade;"},
            { 8501, "&alefsym;"},
            { 8592, "&larr;"},
            { 8593, "&uarr;"},
            { 8594, "&rarr;"},
            { 8595, "&darr;"},
            { 8596, "&harr;"},
            { 8629, "&crarr;"},
            { 8656, "&lArr;"},
            { 8657, "&uArr;"},
            { 8658, "&rArr;"},
            { 8659, "&dArr;"},
            { 8660, "&hArr;"},
            { 8704, "&forall;"},
            { 8706, "&part;"},
            { 8707, "&exist;"},
            { 8709, "&empty;"},
            { 8711, "&nabla;"},
            { 8712, "&isin;"},
            { 8713, "&notin;"},
            { 8715, "&ni;"},
            { 8719, "&prod;"},
            { 8721, "&sum;"},
            { 8722, "&minus;"},
            { 8727, "&lowast;"},
            { 8730, "&radic;"},
            { 8733, "&prop;"},
            { 8734, "&infin;"},
            { 8736, "&ang;"},
            { 8743, "&and;"},
            { 8744, "&or;"},
            { 8745, "&cap;"},
            { 8746, "&cup;"},
            { 8747, "&int;"},
            { 8756, "&there4;"},
            { 8764, "&sim;"},
            { 8773, "&cong;"},
            { 8776, "&asymp;"},
            { 8800, "&ne;"},
            { 8801, "&equiv;"},
            { 8804, "&le;"},
            { 8805, "&ge;"},
            { 8834, "&sub;"},
            { 8835, "&sup;"},
            { 8836, "&nsub;"},
            { 8838, "&sube;"},
            { 8839, "&supe;"},
            { 8853, "&oplus;"},
            { 8855, "&otimes;"},
            { 8869, "&perp;"},
            { 8901, "&sdot;"},
            { 8968, "&lceil;"},
            { 8969, "&rceil;"},
            { 8970, "&lfloor;"},
            { 8971, "&rfloor;"},
            { 9001, "&lang;"},
            { 9002, "&rang;"},
            { 9674, "&loz;"},
            { 9824, "&spades;"},
            { 9827, "&clubs;"},
            { 9829, "&hearts;"},
            { 9830, "&diams;"},
            { 338, "&OElig;"},
            { 339, "&oelig;"},
            { 352, "&Scaron;"},
            { 353, "&scaron;"},
            { 376, "&Yuml;"},
            { 710, "&circ;"},
            { 732, "&tilde;"},
            { 8194, "&ensp;"},
            { 8195, "&emsp;"},
            { 8201, "&thinsp;"},
            { 8204, "&zwnj;"},
            { 8205, "&zwj;"},
            { 8206, "&lrm;"},
            { 8207, "&rlm;"},
            //{ 8211, "&ndash;"},
            //{ 8212, "&mdash;"},
            //{ 8216, "&lsquo;"},
            //{ 8217, "&rsquo;"},
            //{ 8218, "&sbquo;"},
            //{ 8220, "&ldquo;"},
            //{ 8221, "&rdquo;"},
            //{ 8222, "&bdquo;"},
            { 8224, "&dagger;"},
            { 8225, "&Dagger;"},
            { 8240, "&permil;"},
            { 8249, "&lsaquo;"},
            { 8250, "&rsaquo;"},
            { 8364, "&euro;"}
        };

        #endregion

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
        private NotifyIcon notifyIcon1;
        private System.Windows.Forms.MenuItem itmSep1;
        private MenuItem toggleEnabled;
        private ToolTip toolTip1;

		#endregion


		#region Constructors

		public frmMain()
		{
			InitializeComponent();
            notifyIcon1.Visible = true;
		}

		#endregion


		#region Properties - Public



		#endregion

        #region Properties - Private

        private bool enabled = true;        // turns everything on/off

		private bool _copiedRTF = false;	// true if RTF was copied rather than text
        private String _rawClip;        // raw clipboard data
        private String _modifiedClip;   // clipboard text with replacements made

        private bool clipInserting = false; // true when pasting BACK into clipboard

        // an option in the context menu along with its current status
        private struct optionFlag
        {
            public bool active;
            public String text;
            public String shorttext;
            public String desc;
            public MenuItem menuItem;
            public CheckBox buttonToggle;
        }

        // context menu options and app state
        private optionFlag[] CMOptions = new optionFlag[] {
            new optionFlag() { 
                active = true,
                text = "Replace \'smart quotes\'",
                shorttext = "” -> \"",
                desc = "Replaces quote characters added by MS Word, some versions\nof Outlook and others with ASCII equivalent quotes. This replaces\nsingle and double quotes, dashes, bulletpoints and the ellipsis character."
            },
            new optionFlag() { 
                active = false,  
                text = "Keep \'smart quotes\' as entities",
                shorttext = "” -> &&ldquo;",
                desc = "Replaces single and double quotes, dashes, bulletpoints\nand the ellipsis character with HTML entity references.\nUse this to preserve these characters in an ISO-8859-1\nencoded website rather than converting them."
            },
            new optionFlag() { 
                active = true,   
                text = "Replace all named HTML entities",
                shorttext = "Named",
                desc = "Any characters (except for double quotes and HTML brackets)\nwith named HTML entities are converted to those entity codes."
            },
            new optionFlag() { 
                active = true,   
                text = "Convert all high ASCII to numbered entities" ,
                shorttext = "Numbered",
                desc = "Any characters above the standard ASCII range (0x9F) are\nreplaced with HTML entity numbers.\nThis will ensure any strange characters (even Chinese, Korean etc)\nare preserved in an ISO-8859-1 document."
            },
            new optionFlag() { 
                active = false,  
                text = "Convert rich text to plaintext",
                shorttext = "RTF -> Plain",
                desc = "Convert text copied with formatting to be plaintext.\nMost often the text will be converted on pasting into\na plaintext application, but you can force this here if you wish."
            }
        };

        private IntPtr _ClipboardViewerNext;

        #endregion


        #region Methods - Private

        /// <summary>
		/// Register this form as a Clipboard Viewer application
		/// </summary>
		private void RegisterClipboardViewer()
		{
			_ClipboardViewerNext = RAD.ClipMon.Win32.User32.SetClipboardViewer(this.Handle);
		}

		/// <summary>
		/// Remove this form from the Clipboard Viewer list
		/// </summary>
		private void UnregisterClipboardViewer()
		{
            RAD.ClipMon.Win32.User32.ChangeClipboardChain(this.Handle, _ClipboardViewerNext);
		}

		/// <summary>
		/// Search the clipboard contents for configured characters and translate them
		/// </summary>
		/// <param name="iData">The current clipboard contents</param>
		/// <returns>true if characters were replaced, false otherwise</returns>
        private bool replaceQuotes()
        {
            bool found = false;

            _modifiedClip = _rawClip;

            if (CMOptions[4].active)        // convert to plaintext
            {
                _modifiedClip = convertRTFToString(_rawClip);
                found = true;
            }

            if (CMOptions[0].active)        // replace smart quotes
            {
                Dictionary<int, String> replacementsToRun;

                if (CMOptions[1].active)    // replace them with entity refs instead of low-ASCII equivalents
                {
                    replacementsToRun = smartQuotesEscape;
                }
                else
                {
                    replacementsToRun = smartQuotesTranslate;
                }
                if (runReplacements(replacementsToRun))
                {
                    found = true;
                }
            }
            if (CMOptions[2].active)        // replace named entities
            {
                if (runReplacements(namedEntityEscape))
                {
                    found = true;
                }
            }
            if (CMOptions[3].active)        // replace all high ASCII
            {
                if (replaceHighASCII())
                {
                    found = true;
                }
            }

            return found;
        }

        private bool runReplacements(Dictionary<int, String> repls)
        {
            bool found = false;
            foreach (KeyValuePair<int, String> replacement in repls)
            {
                String search = Char.ConvertFromUtf32(replacement.Key);
                String changed = _modifiedClip.Replace(search, replacement.Value);
                if (!found && changed != _modifiedClip)
                {
                    found = true;
                }
                _modifiedClip = changed;
            }
            return found;
        }

        private bool replaceHighASCII()
        {
            bool found = false;
            for (int i = 0; i < _modifiedClip.Length; ++i)
            {
                char c = _modifiedClip[i];
                if (c > 0x9F)
                {
                    found = true;
                    String entityCode = "&#" + (int)(c) + ";";
                    _modifiedClip = _modifiedClip.Remove(i, 1);
                    _modifiedClip = _modifiedClip.Insert(i, entityCode);
                    // we could increment i here but there's no real need since anything we add won't be bad
                }
            }
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
					_copiedRTF = true;

                    contents = convertRTFToString(_rawClip);
				    setNotificationTooltip("RTF copied");
			    }
                else if (iData.GetDataPresent(DataFormats.UnicodeText))
                {
                    _rawClip = (string)iData.GetData(DataFormats.UnicodeText);
					_copiedRTF = false;
					
                    contents = _rawClip;
                    setNotificationTooltip("Text copied");
                }
                if (_rawClip == null)   // nothing was on the clipboard
                {
                    _rawClip = "";
                }
            }
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
            return contents;
		}

        // convert RTF data to plainText using a RichTextBox
        private String convertRTFToString(String str)
        {
            System.Windows.Forms.RichTextBox rtBox = new System.Windows.Forms.RichTextBox();
            try
            {
                rtBox.Rtf = (string)str;
            }
            catch (Exception e)
            {
                return str;     // already a plaintext string
            }
            return rtBox.Text;
        }

        private void setNotificationTooltip(String tt)
        {
            notifyIcon1.Text = tt;
        }

		/// <summary>
		/// Show the clipboard contents in the window 
		/// and show the notification balloon if a link is found
		/// </summary>
		private void ProcessClipboard()
		{
            if (!enabled)
            {
                return;
            }

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

            String strClip = getClipAsString(iData);  // this also assigns _rawClip and _copiedRTF
            
            // show current clipboard text even if there was no match
            if (strClip != "")
            {
                if (_copiedRTF)
                {
                    try
                    {
                        ctlClipboardText.Rtf = _rawClip;   // but show the original RTF formatted string where appropriate
                    }
                    catch (Exception e)
                    {
                        // sometimes this fails when converting from RTF to plaintext and back again
                        ctlClipboardText.Text = _rawClip;
                    }
                }
                else
                {
                    ctlClipboardText.Text = _rawClip;
                }
            }

            if (replaceQuotes())
			{
                bool rtfOutput = _copiedRTF && !CMOptions[4].active;

				// bad quotes were found and have be purged, so update the text with the fixed one
                if (rtfOutput)
                {
					ctlClipboardText.Rtf = _modifiedClip;
				} else {
					ctlClipboardText.Text = _modifiedClip;
				}

				notifyIcon1.ShowBalloonTip(
                    1000, 
                    "Characters fixed", 
                    "Click to view the clipboard contents",
                    ToolTipIcon.Info
                );

                clipInserting = true;
                storeToClipboard(_modifiedClip, rtfOutput);
                clipInserting = false;
			}
		}

        private void storeToClipboard(String value, bool asRTF)
        {
            String err = null;
            IDataObject ido = new DataObject();

            if (asRTF)
            {
                // also set clipboard text property
                String temp = convertRTFToString(_modifiedClip);
                ido.SetData(DataFormats.UnicodeText, true, temp);
                ido.SetData(DataFormats.Rtf, true, _modifiedClip);
            }
            else
            {
                ido.SetData(DataFormats.UnicodeText, true, _modifiedClip);
            }
            try
            {
                Clipboard.SetDataObject(ido, true);
            }
            catch (ExternalException e)
            {
                err = "The clipboard was in use by another application and could not be written to.\nClick to copy the modified text manually.";
            }
            catch (System.Threading.ThreadStateException e)
            {
                err = "The clipboard cannot be pasted to as you are not running this application in single-threaded apartment mode.";
            }
            if (err != null)
            {
                notifyIcon1.ShowBalloonTip(
                    1000,
                    "Character fix failed",
                    err,
                    ToolTipIcon.Error
                );
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
            switch ((RAD.ClipMon.Win32.Msgs)m.Msg)
			{
					//
					// The WM_DRAWCLIPBOARD message is sent to the first window 
					// in the clipboard viewer chain when the content of the 
					// clipboard changes. This enables a clipboard viewer 
					// window to display the new content of the clipboard. 
					//
                case RAD.ClipMon.Win32.Msgs.WM_DRAWCLIPBOARD:
					
					Debug.WriteLine("WindowProc DRAWCLIPBOARD: " + m.Msg, "WndProc");

                    if (!clipInserting)
                    {
                        ProcessClipboard();
                    }

					//
					// Each window that receives the WM_DRAWCLIPBOARD message 
					// must call the SendMessage function to pass the message 
					// on to the next window in the clipboard viewer chain.
					//
                    RAD.ClipMon.Win32.User32.SendMessage(_ClipboardViewerNext, m.Msg, m.WParam, m.LParam);
					break;


					//
					// The WM_CHANGECBCHAIN message is sent to the first window 
					// in the clipboard viewer chain when a window is being 
					// removed from the chain. 
					//
                case RAD.ClipMon.Win32.Msgs.WM_CHANGECBCHAIN:
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
                        RAD.ClipMon.Win32.User32.SendMessage(_ClipboardViewerNext, m.Msg, m.WParam, m.LParam);
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

        private void itmExit_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

		private void itmHide_Click(object sender, System.EventArgs e)
		{
            toggleWindow();
        }

        // toggle replacements on/off
        private void menuItem1_Click(object sender, EventArgs e)
        {
            MenuItem clicked = (System.Windows.Forms.MenuItem)sender;

            enabled = !enabled;
            clicked.Checked = enabled;

            for (int i = 0; i < CMOptions.Length; ++i)
            {
                CMOptions[i].menuItem.Enabled = enabled;
            }
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                toggleWindow();    // toggle window when double clicked
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

        private void clipboardOptionChange(object sender, EventArgs e)
        {
            MenuItem clicked = (System.Windows.Forms.MenuItem)sender;
            for (int i = 0; i < CMOptions.Length; ++i)
            {
                if (CMOptions[i].text == clicked.Text)
                {
                    clicked.Checked = !clicked.Checked;
                    CMOptions[i].active = clicked.Checked;
                    CMOptions[i].buttonToggle.Checked = clicked.Checked;
                    ProcessClipboard();
                }
            }
        }

        // can't be bothered making this generic, to be honest :p
        private void clipboardOptionBtnChange(object sender, EventArgs e)
        {
            CheckBox clicked = (System.Windows.Forms.CheckBox)sender;
            for (int i = 0; i < CMOptions.Length; ++i)
            {
                if (CMOptions[i].shorttext == clicked.Text)
                {
                    CMOptions[i].active = clicked.Checked;
                    CMOptions[i].menuItem.Checked = clicked.Checked;
                    ProcessClipboard();
                }
            }
        }

		#endregion


		#region Event Handlers - Internal

		private void frmMain_Load(object sender, System.EventArgs e)
        {
            // Help for the options
            toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 500;
            toolTip1.ReshowDelay = 500;

            this.cmnuTray.MenuItems.Clear();

            for (int i = 0; i < CMOptions.Length; ++i)
            {
                optionFlag option = CMOptions[i];

                // add system tray menuitem
                MenuItem it = new System.Windows.Forms.MenuItem(option.text);
                it.Checked = option.active;
                it.Click += new EventHandler(clipboardOptionChange);
                this.cmnuTray.MenuItems.Add(it);

                // add form button
                CheckBox btn = new System.Windows.Forms.CheckBox();
                btn.Appearance = System.Windows.Forms.Appearance.Button;
                btn.Location = new System.Drawing.Point(8, 8 + i * 32);
                btn.Name = "btn" + i;
                btn.Size = new System.Drawing.Size(95, 23);
                btn.TabIndex = 1;
                btn.UseVisualStyleBackColor = true;
                btn.Text = option.shorttext;
                toolTip1.SetToolTip(btn, option.desc);
                btn.Checked = option.active;
                btn.Click += new EventHandler(clipboardOptionBtnChange);
                this.Controls.Add(btn);

                CMOptions[i].menuItem = it;
                CMOptions[i].buttonToggle = btn;
            }
            this.cmnuTray.MenuItems.Add(itmSep1);
            this.cmnuTray.MenuItems.Add(itmHide);
            this.cmnuTray.MenuItems.Add(toggleEnabled);
            this.cmnuTray.MenuItems.Add(itmSep2);
            this.cmnuTray.MenuItems.Add(itmExit);

			RegisterClipboardViewer();
		}

		private void frmMain_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			UnregisterClipboardViewer();
		}

        private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
        {
            this.Show();
        }

        private void toggleWindow()
        {
            this.Visible = (!this.Visible);
            itmHide.Text = this.Visible ? "Hide" : "Show";

            if (this.Visible == true)
            {
                if (this.WindowState == FormWindowState.Minimized)
                {
                    this.WindowState = FormWindowState.Normal;
                }
                this.TopMost = true;    // toggling this brings the window to the top of the z-order
                this.TopMost = false;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.menuMain = new System.Windows.Forms.MainMenu(this.components);
            this.cmnuTray = new System.Windows.Forms.ContextMenu();
            this.itmSep1 = new System.Windows.Forms.MenuItem();
            this.itmHide = new System.Windows.Forms.MenuItem();
            this.itmSep2 = new System.Windows.Forms.MenuItem();
            this.itmExit = new System.Windows.Forms.MenuItem();
            this.ctlClipboardText = new System.Windows.Forms.RichTextBox();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.toggleEnabled = new System.Windows.Forms.MenuItem();
            this.SuspendLayout();
            // 
            // cmnuTray
            // 
            this.cmnuTray.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.itmSep1,
            this.itmHide,
            this.toggleEnabled,
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
            this.itmSep2.Index = 3;
            this.itmSep2.Text = "-";
            // 
            // itmExit
            // 
            this.itmExit.Index = 4;
            this.itmExit.MergeOrder = 1000;
            this.itmExit.Text = "E&xit";
            this.itmExit.Click += new System.EventHandler(this.itmExit_Click_1);
            // 
            // ctlClipboardText
            // 
            this.ctlClipboardText.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.ctlClipboardText.DetectUrls = false;
            this.ctlClipboardText.Location = new System.Drawing.Point(109, 0);
            this.ctlClipboardText.Name = "ctlClipboardText";
            this.ctlClipboardText.ReadOnly = true;
            this.ctlClipboardText.Size = new System.Drawing.Size(375, 273);
            this.ctlClipboardText.TabIndex = 0;
            this.ctlClipboardText.Text = "";
            this.ctlClipboardText.WordWrap = false;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.ContextMenu = this.cmnuTray;
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "CP1252 Fixer";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.BalloonTipClicked += new System.EventHandler(this.notifyIcon1_BalloonTipClicked);
            this.notifyIcon1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseClick);
            // 
            // toggleEnabled
            // 
            this.toggleEnabled.Checked = true;
            this.toggleEnabled.Index = 2;
            this.toggleEnabled.Text = "Enabled";
            this.toggleEnabled.Click += new System.EventHandler(this.menuItem1_Click);
            // 
            // frmMain
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
            this.ClientSize = new System.Drawing.Size(484, 273);
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

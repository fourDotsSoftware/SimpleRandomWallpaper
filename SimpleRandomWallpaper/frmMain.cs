using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using Microsoft.Win32;

namespace SimpleRandomWallpaper
{
    public partial class frmMain : SimpleRandomWallpaper.CustomForm
    {
        public static frmMain Instance = null;

        public bool SilentAdd = false;
        public string SilentAddErr = "";        

        public string Err = "";        

        public frmMain()
        {
            InitializeComponent();                        

            Instance = this;

            dt.Columns.Add("foldersep",typeof(bool));
            dt.Columns.Add("subdirs",typeof(bool));
            dt.Columns.Add("filename");           
            dt.Columns.Add("sizekb");
            dt.Columns.Add("fullfilepath");
            dt.Columns.Add("filedate");

            dgFiles.AutoGenerateColumns = false;                        
        }        

        public DataTable dt = new DataTable("table");
        public DataTable dtClipboard = new DataTable("table");

        private bool _IsDirty = false;

        private bool IsDirty
        {
            get { return _IsDirty; }

            set
            {
                _IsDirty = value;

                lblTotal.Text = TranslateHelper.Translate("Total") + " : " + dt.Rows.Count + " " + TranslateHelper.Translate("Images");
            }
        }


        private void tsdbAddFile_ButtonClick(object sender, EventArgs e)
        {
            openFileDialog1.Filter = Module.OpenFilesFilter;
            openFileDialog1.Multiselect = true;

            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SilentAddErr = "";

                try
                {
                    this.Cursor = Cursors.WaitCursor;

                    System.Threading.Thread.Sleep(200);

                    for (int k = 0; k < openFileDialog1.FileNames.Length; k++)
                    {
                        AddFile(openFileDialog1.FileNames[k]);
                        RecentFilesHelper.AddRecentFile(openFileDialog1.FileNames[k]);
                    }
                }
                finally
                {
                    this.Cursor = null;

                    if (SilentAddErr != string.Empty)
                    {
                        frmError f = new frmError(TranslateHelper.Translate("Error"), SilentAddErr);
                        f.ShowDialog(this);
                    }

                    dgFiles_CellClick(null, null);

                    dgFiles.Invalidate();

                    IsDirty = true;
                }
            }
        }

        private void tsdbAddFile_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                SilentAddErr = "";

                AddFile(e.ClickedItem.Text);
                RecentFilesHelper.AddRecentFile(e.ClickedItem.Text);

                IsDirty = true;

            }
            finally
            {
                this.Cursor = null;

                if (SilentAddErr != string.Empty)
                {
                    frmError f = new frmError(TranslateHelper.Translate("Error"), SilentAddErr);
                    f.ShowDialog(this);
                }
            }
        }

        private void tsbRemove_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedCellCollection cells = dgFiles.SelectedCells;
            List<DataGridViewRow> rows = new List<DataGridViewRow>();

            for (int k = 0; k < cells.Count; k++)
            {
                if (rows.IndexOf(dgFiles.Rows[cells[k].RowIndex]) < 0)
                {
                    rows.Add(dgFiles.Rows[cells[k].RowIndex]);
                }
            }

            for (int k = 0; k < rows.Count; k++)
            {
                dgFiles.Rows.Remove(rows[k]);
            }

            dgFiles_CellClick(null, null);

            IsDirty = true;
        }        

        private void tsbClear_Click(object sender, EventArgs e)
        {
            //LockTest();
            //return;

            DialogResult dres = Module.ShowQuestionDialog(TranslateHelper.Translate("Are you sure that you want clear the added files ?"), TranslateHelper.Translate("Clear Added Files ?"));

            if (dres == DialogResult.Yes)
            {
                dt.Rows.Clear();
            }

            IsDirty = true;
        }

        private void tsdbAddFolder_ButtonClick(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = "";
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                SilentAddErr = "";

                AddFolder(folderBrowserDialog1.SelectedPath, false, false);
                RecentFilesHelper.AddRecentFolder(folderBrowserDialog1.SelectedPath);

                IsDirty = true;

                if (SilentAddErr != string.Empty)
                {
                    frmError f = new frmError(TranslateHelper.Translate("Error"), SilentAddErr);
                    f.ShowDialog(this);
                }
            }
        }

        private void tsdbAddFolder_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            SilentAddErr = "";

            AddFolder(e.ClickedItem.Text, false, false);
            RecentFilesHelper.AddRecentFolder(e.ClickedItem.Text);

            IsDirty = true;

            if (SilentAddErr != string.Empty)
            {
                frmError f = new frmError(TranslateHelper.Translate("Error"), SilentAddErr);
                f.ShowDialog(this);
            }
        }

        public void ImportList(string listfilepath)
        {
            string curdir = Environment.CurrentDirectory;

            try
            {
                SilentAdd = true;
                using (StreamReader sr = new StreamReader(listfilepath, Encoding.Default, true))
                {
                    string line = null;

                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.StartsWith("#"))
                        {
                            continue;
                        }

                        string filepath = line;
                        string password = "";

                        try
                        {
                            if (line.StartsWith("\""))
                            {
                                int epos = line.IndexOf("\"", 1);

                                if (epos > 0)
                                {
                                    filepath = line.Substring(1, epos - 1);
                                }
                            }
                            else if (line.StartsWith("'"))
                            {
                                int epos = line.IndexOf("'", 1);

                                if (epos > 0)
                                {
                                    filepath = line.Substring(1, epos - 1);
                                }
                            }

                            int compos = line.IndexOf(",");

                            if (compos > 0)
                            {
                                password = line.Substring(compos + 1);

                                if (!line.StartsWith("\"") && !line.StartsWith("'"))
                                {
                                    filepath = line.Substring(0, compos);
                                }

                                if ((password.StartsWith("\"") && password.EndsWith("\""))
                                    || (password.StartsWith("'") && password.EndsWith("'")))
                                {
                                    if (password.Length == 2)
                                    {
                                        password = "";
                                    }
                                    else
                                    {
                                        password = password.Substring(1, password.Length - 2);
                                    }
                                }

                            }
                        }
                        catch (Exception exq)
                        {
                            SilentAddErr += TranslateHelper.Translate("Error while processing List !") + " " + line + " " + exq.Message + "\r\n";
                        }

                        line = filepath;

                        Environment.CurrentDirectory = System.IO.Path.GetDirectoryName(listfilepath);

                        line = System.IO.Path.GetFullPath(line);

                        if (System.IO.File.Exists(line))
                        {
                            AddFile(line);
                            /*
                            else
                            {
                                SilentAddErr += TranslateHelper.Translate("Error wrong file type !") + " " + line + "\r\n";
                            }*/
                        }
                        else if (System.IO.Directory.Exists(line))
                        {
                            AddFolder(line, false, false);
                        }
                        else
                        {
                            SilentAddErr += TranslateHelper.Translate("Error. File or Directory not found !") + " " + line + "\r\n";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SilentAddErr += TranslateHelper.Translate("Error could not read file !") + " " + ex.Message + "\r\n";
            }
            finally
            {
                Environment.CurrentDirectory = curdir;

                SilentAdd = false;
            }

            IsDirty = true;
        }

        private void tsdbImportList_ButtonClick(object sender, EventArgs e)
        {
            SilentAddErr = "";

            openFileDialog1.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ImportList(openFileDialog1.FileName);
                RecentFilesHelper.ImportListRecent(openFileDialog1.FileName);

                IsDirty = true;

                if (SilentAddErr != string.Empty)
                {
                    frmMessage f = new frmMessage();
                    f.txtMsg.Text = SilentAddErr;
                    f.ShowDialog();

                }
            }
        }

        private void tsdbImportList_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            SilentAddErr = "";

            ImportList(e.ClickedItem.Text);
            RecentFilesHelper.ImportListRecent(e.ClickedItem.Text);

            IsDirty = true;

            if (SilentAddErr != string.Empty)
            {
                frmMessage f = new frmMessage();
                f.txtMsg.Text = SilentAddErr;
                f.ShowDialog();

            }
        }
        /*
        #region Share

        private void tsiFacebook_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareFacebook();
        }

        private void tsiTwitter_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareTwitter();
        }

        private void tsiGooglePlus_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareGooglePlus();
        }

        private void tsiLinkedIn_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareLinkedIn();
        }

        private void tsiEmail_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareEmail();
        }

        #endregion
        */
        public bool AddFile(string filepath)
        {
            string ext = "*" + System.IO.Path.GetExtension(filepath).ToLower() + ";";

            /*
            if (Module.AcceptableMediaInputPattern.IndexOf(ext) < 0)
            {
                SilentAddErr += filepath + "\n\n" + TranslateHelper.Translate("Please add only Word Files !") + "\n\n";

                return false;
            }
            */

            DataRow dr = dt.NewRow();

            FileInfo fi = new FileInfo(filepath);

            long sizekb = fi.Length / 1024;
            dr["filename"] = fi.Name;
            dr["fullfilepath"] = filepath;
            dr["sizekb"] = sizekb.ToString() + "KB";
            dr["filedate"] = fi.LastWriteTime.ToString();                                                

            dt.Rows.Add(dr);

            /*
            if (dt.Rows.Count == 1)
            {
                string outfile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filepath), "mergedDocument.docx");

                RecentFilesHelper.AddRecentOutputFile(outfile);                
            }
            */

            IsDirty = true;

            return true;
        }
        public void AddFolder(string folder_path, bool subdirs,bool silent)
        {            
            if (!silent)
            {
                if (System.IO.Directory.GetDirectories(folder_path).Length > 0)
                {
                    DialogResult dres = Module.ShowQuestionDialog("Would you like to add also Subdirectories ?", TranslateHelper.Translate("Add Subdirectories ?"));

                    if (dres == DialogResult.Yes)
                    {
                        //filez = System.IO.Directory.GetFiles(folder_path, "*.*", SearchOption.AllDirectories);
                        subdirs = true;
                    }
                    else
                    {
                        //filez = System.IO.Directory.GetFiles(folder_path, "*.*", SearchOption.TopDirectoryOnly);
                        subdirs = false;
                    }
                }
                else
                {
                    //filez = System.IO.Directory.GetFiles(folder_path, "*.*", SearchOption.TopDirectoryOnly);
                    subdirs = false;
                }
            }

            try
            {
                this.Cursor = Cursors.WaitCursor;

                DataRow dr = dt.NewRow();

                DirectoryInfo di = new DirectoryInfo(folder_path);

                dr["subdirs"] = subdirs;
                dr["filename"] = System.IO.Path.GetFileName(folder_path);
                dr["fullfilepath"] = folder_path;
                dr["filedate"] = di.LastWriteTime.ToString();
                dr["foldersep"] = true;

                dt.Rows.Add(dr);

            }
            finally
            {
                this.Cursor = null;
            }
        }
        private void SetupOnLoad()
        {
            dgFiles.DataSource = dt;

            //3this.Icon = Properties.Resources.pdf_compress_48;

            this.Text = Module.ApplicationTitle;
            //this.Width = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width;
            //this.Left = 0;
            AddLanguageMenuItems();

            //3DownloadSuggestionsHelper ds = new DownloadSuggestionsHelper();
            //3ds.SetupDownloadMenuItems(downloadToolStripMenuItem);

            AdjustSizeLocation();

            //3SetupOutputFolders();

            //3keepFolderStructureToolStripMenuItem.Checked = Properties.Settings.Default.KeepFolderStructure;

            RecentFilesHelper.FillMenuRecentFile();
            RecentFilesHelper.FillMenuRecentFolder();
            RecentFilesHelper.FillMenuRecentImportList();
            //3RecentFilesHelper.FillRecentOutputFile(); 

            string enabled = RegistryHelper2.GetKeyValue(Module.ApplicationName, "Enabled");

            if (enabled==bool.TrueString)
            {
                chkEnableCurrentUser.Checked = true;
            }

            string interval = RegistryHelper2.GetKeyValue(Module.ApplicationName, "Interval");

            string enableforall = RegistryHelper2.GetKeyValueLMLowPriv(Module.ApplicationName, "EnabledForAll");

            if (enableforall == bool.TrueString)
            {
                chkEnableForAllUsers.Checked = true;

                interval = RegistryHelper2.GetKeyValueLMLowPriv(Module.ApplicationName, "Interval");
            }

            if (!Properties.Settings.Default.Initialized)
            {
                if (!chkEnableCurrentUser.Checked && !chkEnableForAllUsers.Checked)
                {
                    chkEnableCurrentUser.Checked = true;
                }
            }

            if (interval == string.Empty)
            {
                nudD.Value = 1;
                nudM.Value = 0;
                nudS.Value = 0;
                nudH.Value = 0;
            }
            else
            {
                int intervalSecs = int.Parse(interval);

                TimeSpan ts = new TimeSpan(0, 0, intervalSecs);

                nudD.Value = ts.Days;
                nudH.Value = ts.Hours;
                nudM.Value = ts.Minutes;
                nudS.Value = ts.Seconds;                
            }

            /*
            string file = "";

            string sid = WindowsIdentity.GetCurrent().User.ToString();

            if (chkEnableCurrentUser.Checked)
            {
                file = System.IO.Path.Combine(WallpaperChanger.SettingsFolder, sid + ".images.txt");
            }
            else if (chkEnableForAllUsers.Checked)
            {
                file = System.IO.Path.Combine(WallpaperChanger.SettingsFolder, "allusers.images.txt");
            }
            else
            {
                return;
            }

            if (System.IO.File.Exists(file))
            {
                using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    using (StreamReader sr = new StreamReader(fs))
                    {
                        string line = null;

                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line == "subdirs" || line == "nosubdirs")
                            {

                            }
                            else if (System.IO.File.Exists(line) && Module.IsAcceptableMediaInput(line))
                            {
                                AddFile(line);
                            }
                            else if (System.IO.Directory.Exists(line))
                            {
                                AddFolder(line);
                            }
                        }
                    }
                }
            }
            */

            List<string> lstimg = new List<string>();            

            RegistryKey regkey = Registry.CurrentUser;

            if (chkEnableForAllUsers.Checked)
            {
                regkey = Registry.LocalMachine;
            }

            RegistryKey regkey2 = regkey.OpenSubKey("Software", false);

            if (regkey2 == null) return;

            regkey = regkey2.OpenSubKey("4dots Software", false);

            if (regkey == null) return;

            regkey2 = regkey.OpenSubKey(Module.ApplicationName, false);

            if (regkey2 == null) return;

            regkey = regkey2.OpenSubKey("Images", false);

            if (regkey == null) return;

            string[] valnames = regkey.GetValueNames();

            for (int k = 0; k < valnames.Length; k++)
            {
                string line = regkey.GetValue(valnames[k]).ToString();

                bool subdirs = false;

                if (line.StartsWith("|||SUBDIRS|||"))
                {
                    subdirs = true;

                    line = line.Substring("|||SUBDIRS|||".Length);
                }

                if (System.IO.File.Exists(line))
                {
                    AddFile(line);
                }
                else if (System.IO.Directory.Exists(line))
                {
                    AddFolder(line, subdirs,true);                    
                }
            }

            dgFiles_CellClick(null, null);

            checkForNewVersionEachWeekToolStripMenuItem.Checked = Properties.Settings.Default.CheckWeek;
        }
        private void AdjustSizeLocation()
        {
            if (Properties.Settings.Default.Maximized)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                if (Properties.Settings.Default.Width == -1)
                {
                    this.CenterToScreen();
                    return;
                }
                else
                {
                    this.Width = Properties.Settings.Default.Width;
                }
                if (Properties.Settings.Default.Height != -1)
                {
                    this.Height = Properties.Settings.Default.Height;
                }

                if (Properties.Settings.Default.Left != -1)
                {
                    this.Left = Properties.Settings.Default.Left;
                }

                if (Properties.Settings.Default.Top != -1)
                {
                    this.Top = Properties.Settings.Default.Top;
                }

                if (this.Width < 300)
                {
                    this.Width = 300;
                }

                if (this.Height < 300)
                {
                    this.Height = 300;
                }

                if (this.Left < 0)
                {
                    this.Left = 0;
                }

                if (this.Top < 0)
                {
                    this.Top = 0;
                }
            }

        }

        private void SaveSizeLocation()
        {
            Properties.Settings.Default.Maximized = (this.WindowState == FormWindowState.Maximized);
            Properties.Settings.Default.Left = this.Left;
            Properties.Settings.Default.Top = this.Top;
            Properties.Settings.Default.Width = this.Width;
            Properties.Settings.Default.Height = this.Height;
            Properties.Settings.Default.Save();

        }

        #region Localization

        private void AddLanguageMenuItems()
        {
            for (int k = 0; k < frmLanguage.LangCodes.Count; k++)
            {
                ToolStripMenuItem ti = new ToolStripMenuItem();
                ti.Text = frmLanguage.LangDesc[k];
                ti.Tag = frmLanguage.LangCodes[k];
                ti.Image = frmLanguage.LangImg[k];

                if (Properties.Settings.Default.Language == frmLanguage.LangCodes[k])
                {
                    ti.Checked = true;
                }

                ti.Click += new EventHandler(tiLang_Click);

                if (k < 25)
                {
                    languages1ToolStripMenuItem.DropDownItems.Add(ti);
                }
                else
                {
                    languages2ToolStripMenuItem.DropDownItems.Add(ti);
                }

                //languageToolStripMenuItem.DropDownItems.Add(ti);
            }
        }

        void tiLang_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ti = (ToolStripMenuItem)sender;
            string langcode = ti.Tag.ToString();
            ChangeLanguage(langcode);

            //for (int k = 0; k < languageToolStripMenuItem.DropDownItems.Count; k++)
            for (int k = 0; k < languages1ToolStripMenuItem.DropDownItems.Count; k++)
            {
                ToolStripMenuItem til = (ToolStripMenuItem)languages1ToolStripMenuItem.DropDownItems[k];
                if (til == ti)
                {
                    til.Checked = true;
                }
                else
                {
                    til.Checked = false;
                }
            }

            for (int k = 0; k < languages2ToolStripMenuItem.DropDownItems.Count; k++)
            {
                ToolStripMenuItem til = (ToolStripMenuItem)languages2ToolStripMenuItem.DropDownItems[k];
                if (til == ti)
                {
                    til.Checked = true;
                }
                else
                {
                    til.Checked = false;
                }
            }
        }

        private bool InChangeLanguage = false;

        private void ChangeLanguage(string language_code)
        {
            try
            {
                InChangeLanguage = true;

                Properties.Settings.Default.Language = language_code;
                frmLanguage.SetLanguage();

                Properties.Settings.Default.Save();
                Module.ShowMessage("Please restart the application !");
                Application.Exit();
                return;

                bool maximized = (this.WindowState == FormWindowState.Maximized);
                this.WindowState = FormWindowState.Normal;

                /*
                RegistryKey key = Registry.CurrentUser;
                RegistryKey key2 = Registry.CurrentUser;

                try
                {
                    key = key.OpenSubKey("Software\\4dots Software", true);

                    if (key == null)
                    {
                        key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\4dots Software");
                    }

                    key2 = key.OpenSubKey(frmLanguage.RegKeyName, true);

                    if (key2 == null)
                    {
                        key2 = key.CreateSubKey(frmLanguage.RegKeyName);
                    }

                    key = key2;

                    //key.SetValue("Language", language_code);
                    key.SetValue("Menu Item Caption", TranslateHelper.Translate("Change PDF Properties"));
                }
                catch (Exception ex)
                {
                    Module.ShowError(ex);
                    return;
                }
                finally
                {
                    key.Close();
                    key2.Close();
                }
                */
                //1SaveSizeLocation();

                //3SavePositionSize();

                this.Controls.Clear();

                InitializeComponent();

                SetupOnLoad();

                if (maximized)
                {
                    this.WindowState = FormWindowState.Maximized;
                }

                this.ResumeLayout(true);
            }
            finally
            {
                InChangeLanguage = false;
            }
        }

        #endregion        

        private void frmMain_Load(object sender, EventArgs e)
        {            
            SetupOnLoad();

            if (!Module.IsFromWindowsExplorer && !Module.IsCommandLine && Properties.Settings.Default.CheckWeek)
            {
                UpdateHelper.InitializeCheckVersionWeek();
            }

            /*
            if (Module.args != null)
            {
                AddVisual(Module.args);
            }
            */

            //CompressZIPPackage();

            ResizeControls();
        }

        private void AddVisual(string[] argsvisual)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                //Module.ShowMessage("Is From Windows Explorer");                                

                for (int k = 0; k < argsvisual.Length; k++)
                {
                    if (System.IO.File.Exists(argsvisual[k]))
                    {
                        AddFile(argsvisual[k]);

                    }
                    else if (System.IO.Directory.Exists(argsvisual[k]))
                    {
                        AddFolder(argsvisual[k], false, false);
                    }
                }
            }
            finally
            {
                this.Cursor = null;
            }
        }


        #region Help

        private void helpGuideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //System.Diagnostics.Process.Start(Application.StartupPath + "\\Video Cutter Joiner Expert - User's Manual.chm");
            System.Diagnostics.Process.Start(Module.HelpURL);
        }

        private void pleaseDonateToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.4dots-software.com/donate.php");
        }

        private void dotsSoftwarePRODUCTCATALOGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.4dots-software.com/downloads/4dots-Software-PRODUCT-CATALOG.pdf");
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmAbout f = new frmAbout();
            f.ShowDialog();
        }

        private void tiHelpFeedback_Click(object sender, EventArgs e)
        {
            /*
            frmUninstallQuestionnaire f = new frmUninstallQuestionnaire(false);
            f.ShowDialog();
            */

            System.Diagnostics.Process.Start("https://www.4dots-software.com/support/bugfeature.php?app=" + System.Web.HttpUtility.UrlEncode(Module.ShortApplicationTitle));
        }

        private void followUsOnTwitterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("http://www.twitter.com/4dotsSoftware");
        }

        private void visit4dotsSoftwareWebsiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.4dots-software.com");
        }

        private void checkForNewVersionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateHelper.CheckVersion(false);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Application.Exit();
            }
            catch { }
        }

        #endregion

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSizeLocation();

            Properties.Settings.Default.CheckWeek = checkForNewVersionEachWeekToolStripMenuItem.Checked;

            Properties.Settings.Default.Save();
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int k = 0; k < dgFiles.Rows.Count; k++)
            {
                dgFiles.Rows[k].Selected = true;
            }
        }

        private void seelctNoneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int k = 0; k < dgFiles.Rows.Count; k++)
            {
                dgFiles.Rows[k].Selected = false;
            }
        }

        private void invertSelectionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int k = 0; k < dgFiles.Rows.Count; k++)
            {
                dgFiles.Rows[k].Selected = !dgFiles.Rows[k].Selected;
            }
        }

        #region Grid Context menu

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgFiles.CurrentRow == null) return;

            DataRowView drv = (DataRowView)dgFiles.CurrentRow.DataBoundItem;

            DataRow dr = drv.Row;

            string filepath = dr["fullfilepath"].ToString();

            System.Diagnostics.Process.Start(filepath);
        }

        private void exploreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRowView drv = (DataRowView)dgFiles.CurrentRow.DataBoundItem;

            DataRow dr = drv.Row;

            string filepath = dr["fullfilepath"].ToString();

            string args = string.Format("/e, /select, \"{0}\"", filepath);

            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "explorer";
            info.UseShellExecute = true;
            info.Arguments = args;
            Process.Start(info);
        }

        private void copyFullFilePathToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRowView drv = (DataRowView)dgFiles.CurrentRow.DataBoundItem;

            DataRow dr = drv.Row;

            string filepath = dr["fullfilepath"].ToString();

            Clipboard.Clear();

            Clipboard.SetText(filepath);
        }

        private void cmsFiles_Opening(object sender, CancelEventArgs e)
        {
            Point p = dgFiles.PointToClient(new Point(Control.MousePosition.X, Control.MousePosition.Y));
            DataGridView.HitTestInfo hit = dgFiles.HitTest(p.X, p.Y);

            if (hit.Type == DataGridViewHitTestType.Cell)
            {
                dgFiles.CurrentCell = dgFiles.Rows[hit.RowIndex].Cells[hit.ColumnIndex];
            }

            if (dgFiles.CurrentRow == null)
            {
                e.Cancel = true;
            }
        }
        #endregion

        #region Drag and Drop

        private void dgFiles_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void dgFiles_DragOver(object sender, DragEventArgs e)
        {
            if ((e.AllowedEffect & DragDropEffects.Copy) == DragDropEffects.Copy)
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void dgFiles_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                string[] filez = (string[])e.Data.GetData(DataFormats.FileDrop);

                for (int k = 0; k < filez.Length; k++)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;

                        if (System.IO.File.Exists(filez[k]))
                        {
                            AddFile(filez[k]);
                        }
                        else if (System.IO.Directory.Exists(filez[k]))
                        {
                            AddFolder(filez[k], false, false);
                        }
                    }
                    finally
                    {
                        this.Cursor = null;
                    }
                }
            }
        }

        #endregion        

        private void saveDocumentsListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog opf = new SaveFileDialog();
            opf.Filter = "Text Files (*.txt)|*.txt";

            if (opf.ShowDialog() == DialogResult.OK)
            {
                using (System.IO.StreamWriter sw = new StreamWriter(opf.FileName))
                {
                    for (int k = 0; k < dt.Rows.Count; k++)
                    {
                        sw.WriteLine("\""+dt.Rows[k]["fullfilepath"].ToString()+"\"");
                    }
                }
            }
        }                       

        private void dgFiles_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgFiles.Columns[e.ColumnIndex].Name == "colFolder")
            {
                DataRowView dv = (DataRowView)dgFiles.Rows[e.RowIndex].DataBoundItem;

                if (dv.Row["foldersep"].ToString() == bool.TrueString)
                {
                    e.Value = Properties.Resources.folder;
                }
                else
                {
                    e.Value = new System.Drawing.Bitmap(1, 1);
                }
            }
        }

        private void dgFiles_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgFiles.CurrentRow==null)
            {
                picImage.Image = null;
            }
            else
            {
                try
                {
                    string fp = dgFiles.CurrentRow.Cells["colFullFilePath"].Value.ToString();

                    if (System.IO.File.Exists(fp))
                    {
                        Image img = Image.FromFile(fp);

                        picImage.Image = img;
                    }
                    else
                    {
                        picImage.Image = null;
                    }
                }
                catch
                {
                    picImage.Image = null;
                }
            }
        }

        private void nudD_ValueChanged(object sender, EventArgs e)
        {            
            TimeSpan ts = new TimeSpan((int)nudD.Value, (int)nudH.Value, (int)nudM.Value, (int)nudS.Value);

            RegistryHelper2.SetKeyValue(Module.ApplicationName, "Interval", ts.TotalSeconds.ToString());
        }

        private void chkDisable_CheckedChanged(object sender, EventArgs e)
        {
            RegistryHelper2.SetKeyValue(Module.ApplicationName, "Disabled",chkEnableForAllUsers.Checked.ToString());
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                this.Enabled = false;

                string file = System.IO.Path.GetTempFileName();

                List<string> lstimg = new List<string>();

                using (StreamWriter sw = new StreamWriter(file, false))
                {
                    for (int k = 0; k < dt.Rows.Count; k++)
                    {
                        if (System.IO.File.Exists(dt.Rows[k]["fullfilepath"].ToString()))
                        {
                            sw.WriteLine(dt.Rows[k]["fullfilepath"].ToString());
                            lstimg.Add(dt.Rows[k]["fullfilepath"].ToString());
                }
                        else if (System.IO.Directory.Exists(dt.Rows[k]["fullfilepath"].ToString()))
                        {
                            bool subdirs = dt.Rows[k]["subdirs"].ToString() == bool.TrueString;

                            string strsubdirs = "";

                            if (subdirs)
                            {
                                strsubdirs = "|||SUBDIRS|||";
                            }

                            //sw.WriteLine(subdirs ? "subdirs" : "nosubdirs");

                            sw.WriteLine(strsubdirs+dt.Rows[k]["fullfilepath"].ToString());

                            lstimg.Add(strsubdirs+dt.Rows[k]["fullfilepath"].ToString());                                                     
                        }
                    }
                }

                TimeSpan ts = new TimeSpan((int)nudD.Value, (int)nudH.Value, (int)nudM.Value, (int)nudS.Value);

                string args = "-SimpleRandomWallpaper \"" + Module.ApplicationName + "\" \"EnabledForAll\" \"" + chkEnableForAllUsers.Checked.ToString() + "\" \"Interval\" \"" + ts.TotalSeconds.ToString() + "\" \""+file+"\"";

                bool suc = Module.RunAdminAction(args);

                if (!suc)
                {
                    this.Enabled = true;

                    Module.ShowMessage("Error could not set Settings !");

                    return;
                }

                RegistryHelper2.SetKeyValue(Module.ApplicationName, "Enabled", chkEnableCurrentUser.Checked.ToString());

                RegistryHelper2.SetKeyValue(Module.ApplicationName, "Interval", ts.TotalSeconds.ToString());

                RegistryHelper2.RegisterKeyFolder(Module.ApplicationName, "Images", lstimg);

                //string file = "";

                /*
                if (!System.IO.Directory.Exists(WallpaperChanger.SettingsFolder))
                {
                    System.IO.Directory.CreateDirectory(WallpaperChanger.SettingsFolder);
                }

                string sid = WindowsIdentity.GetCurrent().User.ToString();
                */

                /*
                bool setallusers = false;

                if (chkEnableCurrentUser.Checked)
                {
                    file = System.IO.Path.Combine(WallpaperChanger.SettingsFolder, sid + ".images.txt");
                }
                else if (chkEnableForAllUsers.Checked)
                {
                    file = System.IO.Path.Combine(WallpaperChanger.SettingsFolder, "allusers.images.txt");

                    /*
                    if (chkEnableForAllUsers.Checked && System.IO.File.Exists(file))
                    {
                        setallusers = true;

                        args = "-fileaccess \"" + System.IO.Path.Combine(WallpaperChanger.SettingsFolder, "allusers.images.txt") + "\"";

                        suc = Module.RunAdminAction(args);

                        if (!suc)
                        {
                            this.Enabled = true;

                            Module.ShowMessage("Error could not set Settings !");

                            return;
                        }
                    }*/
                    /*
                }
                */                

                /*
                if (!setallusers && chkEnableForAllUsers.Checked && System.IO.File.Exists(file))
                {
                    setallusers = true;

                    args = "-fileaccess \"" + System.IO.Path.Combine(WallpaperChanger.SettingsFolder, "allusers.images.txt") + "\"";

                    suc = Module.RunAdminAction(args);

                    if (!suc)
                    {
                        this.Enabled = true;

                        Module.ShowMessage("Error could not set Settings !");

                        return;
                    }
                }*/
            }
            finally
            {
                this.Enabled = true;
            }

            Properties.Settings.Default.Initialized = true;
            Properties.Settings.Default.Save();

            Application.Exit();
        }

        private void importListFromExcelFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel Files (*.xls;*.xlsx;*.xlt)|*.xls;*.xlsx;*.xlt";
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ExcelImporter xl = new ExcelImporter();
                xl.ImportListExcel(openFileDialog1.FileName);                
            }
        }

        private void enterListOfFIlesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string txt = "";

            for (int k = 0; k < dt.Rows.Count; k++)
            {
                txt += dt.Rows[k]["fullfilepath"].ToString() + "\r\n";
            }

            frmMultipleFiles f = new frmMultipleFiles(txt);

            if (f.ShowDialog() == DialogResult.OK)
            {
                dt.Rows.Clear();

                for (int k = 0; k < f.txtFiles.Lines.Length; k++)
                {
                    if (System.IO.File.Exists(f.txtFiles.Lines[k]))
                    {
                        AddFile(f.txtFiles.Lines[k]);
                    }
                    else if (System.IO.Directory.Exists(f.txtFiles.Lines[k]))
                    {
                        AddFolder(f.txtFiles.Lines[k], false, false);
                    }
                }
            }
        }
    }
}

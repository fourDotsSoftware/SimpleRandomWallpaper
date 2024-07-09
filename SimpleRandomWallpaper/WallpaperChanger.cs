using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Win32;

namespace SimpleRandomWallpaper
{
    class WallpaperChanger
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern Int32 SystemParametersInfo(
    UInt32 uiAction, UInt32 uiParam, String pvParam, UInt32 fWinIni);

        private static UInt32 SPI_SETDESKWALLPAPER = 20;

        private static UInt32 SPIF_UPDATEINIFILE = 0x1;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindow(
         [MarshalAs(UnmanagedType.LPTStr)] string lpClassName,
         [MarshalAs(UnmanagedType.LPTStr)] string lpWindowName);
        [DllImport("user32.dll")]
        public static extern IntPtr SetParent(
         IntPtr hWndChild,      // handle to window
         IntPtr hWndNewParent   // new parent window
         );

        [DllImport("user32.dll")]
        static extern IntPtr GetDesktopWindow();

        public static bool SetWallpaper(string file)
        {
            try
            {
                string tpath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Module.ApplicationName);

                if (!System.IO.Directory.Exists(tpath))
                {
                    System.IO.Directory.CreateDirectory(tpath);
                }

                string[] filez = System.IO.Directory.GetFiles(tpath);

                for (int k = 0; k < filez.Length; k++)
                {
                    try
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(filez[k]);
                        fi.Attributes = System.IO.FileAttributes.Normal;
                        fi.Delete();
                    }
                    catch { }
                }

                string npath = System.IO.Path.Combine(tpath, Guid.NewGuid().ToString() + ".bmp");

                System.Drawing.Image img = System.Drawing.Image.FromFile(file);

                img.Save(npath, System.Drawing.Imaging.ImageFormat.Bmp);

                SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, npath, SPIF_UPDATEINIFILE);
            }
            catch
            {

            }



            return true;
        }

        public static string SettingsFolder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + Module.ApplicationName;
        public static void SetRandomWallpaper(string sid)
        {
            //string settingsFile = System.IO.Path.Combine(SettingsFolder, sid + ".images.txt");

            //if (System.IO.File.Exists(settingsFile))
            //{
            List<string> lstimg = new List<string>();

            //using (FileStream fs = new FileStream(settingsFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            //{
            /*
                using (StreamReader sr = new StreamReader(fs))
                {
                    string line = null;

                    while ((line = sr.ReadLine()) != null)
                    {*/

            RegistryKey regkey = Registry.CurrentUser;

            if (sid == "allusers")
            {
                regkey = Registry.LocalMachine;
            }

            RegistryKey regkey2 = regkey.OpenSubKey("Software", false);

            if (regkey2 == null) return;

            regkey = regkey2.OpenSubKey("softpcapps Software", false);

            if (regkey == null) return;

            regkey2 = regkey.OpenSubKey(Module.ApplicationName, false);

            if (regkey2 == null) return;

            regkey = regkey2.OpenSubKey("Images", false);

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
                    lstimg.Add(line);
                }
                else if (System.IO.Directory.Exists(line))
                {
                    string[] filez = null;

                    if (!subdirs)
                    {
                        filez = System.IO.Directory.GetFiles(line, "*.*", System.IO.SearchOption.TopDirectoryOnly);
                    }
                    else
                    {
                        filez = System.IO.Directory.GetFiles(line, "*.*", System.IO.SearchOption.AllDirectories);
                    }

                    for (int m = 0; m < filez.Length; m++)
                    {
                        if (Module.IsAcceptableMediaInput(filez[m]))
                        {
                            lstimg.Add(filez[m]);
                        }
                    }
                }
            }


            if (lstimg.Count == 0)
            {
                return;
            }

            Random r = new Random();
            int l = r.Next(0, lstimg.Count - 1);

            string wfile = lstimg[l];

            SetWallpaper(wfile);

            RegistryHelper2.SetKeyValue(Module.ApplicationName, "LastChange", DateTime.Now.ToFileTime().ToString());
        }
    }
}
    


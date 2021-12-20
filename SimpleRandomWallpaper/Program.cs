using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SimpleRandomWallpaper
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args != null && args.Length > 0)
            {
                string str = "";
                for (int k = 0; k < args.Length; k++)
                {
                    str += args[k] + " ";

                    args[k] = RemoveQuotes(args[k]);
                }

                //Module.ShowMessage(str);
            }
            /*
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("SimpleRandomWallpaper");

            for (int k=0;k<procs.Length;k++)
            {
                try
                {
                    if (procs[k].Id != System.Diagnostics.Process.GetCurrentProcess().Id)
                    {
                        procs[k].Kill();
                    }
                }
                catch { }
            }
            */

            ExceptionHandlersHelper.AddUnhandledExceptionHandlers();

            if ((args != null) && (args.Length>0) && (args[0] == "-setpermissions"))
            {
                if (!System.IO.Directory.Exists(WallpaperChanger.SettingsFolder))
                {
                    System.IO.Directory.CreateDirectory(WallpaperChanger.SettingsFolder);
                }

                Module.RunAdminAction("-setpermissions \"" + WallpaperChanger.SettingsFolder + "\"");

                return;

            }

            if (args!=null && args.Length>0 && (args[0].ToLower().Trim()=="-setwallpaper"))
            {
                string sid = args[1].Trim();

                WallpaperChanger.SetRandomWallpaper(sid);

                return;
            }


            frmLanguage.SetLanguages();
            frmLanguage.SetLanguage();

            if (args.Length > 0 && args.Length > 0 && args[0].StartsWith("/uninstall"))
            {
                Module.DeleteApplicationSettingsFile();

                System.Diagnostics.Process.Start("https://www.4dots-software.com/support/bugfeature.php?uninstall=true&app=" + System.Web.HttpUtility.UrlEncode(Module.ShortApplicationTitle));

                Environment.Exit(0);

                return;
            }

            Module.args = args;

            Application.Run(new frmMain());
        }

        private static string RemoveQuotes(string str)
        {
            if ((str.StartsWith("\"") && str.EndsWith("\"")) ||
                    (str.StartsWith("'") && str.EndsWith("'")))
            {
                if (str.Length > 2)
                {
                    str = str.Substring(1, str.Length - 2);
                }
                else
                {
                    str = "";
                }
            }

            return str;
        }
    }
}

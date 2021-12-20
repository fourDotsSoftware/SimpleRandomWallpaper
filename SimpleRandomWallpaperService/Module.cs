using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SimpleRandomWallpaperService
{
    class Module
    {
        public static string ApplicationName = "Simple Random Wallpaper";

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int wMsg, int wParam,
        int lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool LockWindowUpdate(IntPtr hWndLock);              
        
        public static bool IsCommandLine = false;
        public static bool IsFromWindowsExplorer = false;
                
        public static string UserDocumentsFolder = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\4dots Free PDF Compress";                                 

        [DllImport("shell32.dll")]
		public static extern Int32 SHParseDisplayName(
			[MarshalAs(UnmanagedType.LPWStr)]
			String pszName,
			IntPtr pbc,
			out IntPtr ppidl,
			UInt32 sfgaoIn,
			out UInt32 psfgaoOut);

        [DllImport("shell32.dll", ExactSpelling=true, SetLastError=true, CharSet=CharSet.Unicode)]
        public static extern Int32 SHOpenFolderAndSelectItems(IntPtr pidlFolder , UInt32 cidl, [MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl , UInt32 dwFlags);

        /*
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int wMsg, int wParam,
        int lParam);

        [DllImport("user32.dll")]
        public static extern bool LockWindowUpdate(IntPtr hWndLock);
        */               

        public static bool IsDebugMode
        {
            get
            {
                return System.IO.Path.GetFileName(Application.StartupPath).ToLower() == "debug";
                
            }
        }
        
        public static string GetRelativePath(string mainDirPath, string absoluteFilePath)
        {
            string[] firstPathParts = mainDirPath.Trim(Path.DirectorySeparatorChar).Split(Path.DirectorySeparatorChar);
            string[] secondPathParts = absoluteFilePath.Trim(Path.DirectorySeparatorChar).Split(Path.DirectorySeparatorChar);

            int sameCounter = 0;
            for (int i = 0; i < Math.Min(firstPathParts.Length,
            secondPathParts.Length); i++)
            {
                if (
                !firstPathParts[i].ToLower().Equals(secondPathParts[i].ToLower()))
                {
                    break;
                }
                sameCounter++;
            }

            if (sameCounter == 0)
            {
                return absoluteFilePath;
            }

            string newPath = String.Empty;
            for (int i = sameCounter; i < firstPathParts.Length; i++)
            {
                if (i > sameCounter)
                {
                    newPath += Path.DirectorySeparatorChar;
                }
                newPath += "..";
            }
            if (newPath.Length == 0)
            {
                newPath = ".";
            }
            for (int i = sameCounter; i < secondPathParts.Length; i++)
            {
                newPath += Path.DirectorySeparatorChar;
                newPath += secondPathParts[i];
            }
            return newPath;
        }

        // Useful constants
        public static IntPtr WTS_CURRENT_SERVER_HANDLE = IntPtr.Zero;
        public static int WTS_CURRENT_SESSION = -1;

        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSSendMessage(
                IntPtr hServer,
                [MarshalAs(UnmanagedType.I4)] int SessionId,
                String pTitle,
                [MarshalAs(UnmanagedType.U4)] int TitleLength,
                String pMessage,
                [MarshalAs(UnmanagedType.U4)] int MessageLength,
                [MarshalAs(UnmanagedType.U4)] int Style,
                [MarshalAs(UnmanagedType.U4)] int Timeout,
                [MarshalAs(UnmanagedType.U4)] out int pResponse,
                bool bWait);


        public static void ShowMessageService(string smsg)
        {
            String msg = (String)smsg;

            bool result = false;
            String title = "Alert";
            int tlen = title.Length;
            //String msg = sb.ToString();
            int mlen = msg.Length;
            int resp = 7;
            result = WTSSendMessage(WTS_CURRENT_SERVER_HANDLE, WTS_CURRENT_SESSION, title, tlen, msg, mlen, 4, 3, out resp, true);
            int err = Marshal.GetLastWin32Error();
        }

        public static void ShowMessage(string msg)
        {
            if (msg == string.Empty) return;

            if (Module.IsCommandLine)
            {
                Console.WriteLine(msg);
            }
            else
            {
                /*
                //MessageBox.Show(TranslateHelper.Translate(msg),ApplicationName,MessageBoxButtons.OK,MessageBoxIcon.Information,MessageBoxDefaultButton.Button1,MessageBoxOptions.ServiceNotification);
                //ShowMessageService(msg);
                if (MainHelper.Instance != null)
                {
                    MainHelper.Instance.sw.WriteLine(msg);
                }*/
            }
        }

        public static DialogResult ShowQuestionDialog(string msg, string caption)
        {
            return MessageBox.Show(TranslateHelper.Translate(msg), TranslateHelper.Translate(caption), MessageBoxButtons.YesNo, MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
        }


        public static void ShowError(Exception ex)
        {
            ShowError("Error", ex);
        }

        public static void ShowError(string msg)
        {
            if (Module.IsCommandLine)
            {
                Console.WriteLine("Error:" + msg);
            }
            else
            {
                try
                {
                    //MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //MessageBox.Show(TranslateHelper.Translate(msg), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                    //ShowMessageService(msg);

                    /*
                    if (MainHelper.Instance != null)
                    {
                        MainHelper.Instance.sw.WriteLine(msg);
                    }*/
                }
                catch { }
            }

        }


        public static void ShowError(string msg, Exception ex)
        {
            ShowError(msg + "\n\n" + ex.Message);
        }

        public static void ShowError(string msg, string exstr)
        {
            ShowError(msg + "\n\n" + exstr);
        }

        public static DialogResult ShowQuestionDialogYesFocus(string msg, string caption)
        {
            return MessageBox.Show(TranslateHelper.Translate(msg), TranslateHelper.Translate(caption), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult ShowQuestionDialogWithCancelYesFocus(string msg, string caption)
        {
            return MessageBox.Show(TranslateHelper.Translate(msg), TranslateHelper.Translate(caption), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
        }

        public static void CurrentDomain_ProcessExit(object sender, EventArgs e)
        {           
            
        }
                
        public static int _Modex64 = -1;

        public static bool Modex64
        {
            get
            {
                if (_Modex64 == -1)
                {
                    if (Marshal.SizeOf(typeof(IntPtr)) == 8)
                    {
                        _Modex64 = 1;
                        return true;
                    }
                    else
                    {
                        _Modex64 = 0;
                        return false;
                    }
                }
                else if (_Modex64 == 1)
                {
                    return true;
                }
                else if (_Modex64 == 0)
                {
                    return false;
                }
                return false;
            }
        }                             
        public static bool IsLegalFilename(string name)
        {
            try
            {
                System.IO.FileInfo fileInfo = new System.IO.FileInfo(name);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static string DecimalToString(decimal dec)
        {
            return DecimalToString(dec, 1);
        }

        public static string DecimalToString(decimal dec,int decimal_places)
        {
            string format="#0";
            
            if (decimal_places>0)
            {
                format+=".";
            }

            for (int k=0;k<decimal_places;k++)
            {
                format+="0";
            }
            
             return dec.ToString(format,new System.Globalization.CultureInfo("en-US")).Replace(",", ".");
        }

        public static decimal StringToDecimal(string str)
        {
            if (string.IsNullOrEmpty(str)) return 0;

            int epos = str.LastIndexOf(".");

            if (epos < 0)
            {
                epos = str.LastIndexOf(",");
            }

            if (epos < 0)
            {
                bool ihask = false;

                string sintegral = str;

                if (sintegral.ToLower().IndexOf("k") >= 0)
                {
                    ihask = true;
                }

                int integral_part = int.Parse(sintegral.ToLower().Replace("k", ""));

                return (decimal)integral_part;
            }
            else
            {
                bool ihask = false;

                string sintegral = str.Substring(0,epos);

                if (str.ToLower().IndexOf("k") >= 0)
                {
                    ihask = true;
                }

                int integral_part = int.Parse(sintegral.ToLower().Replace("k",""));

                

                string sdecimal = str.Substring(epos + 1, str.Length - epos - 1).ToLower().Replace("k", "");
                
                int decimal_part=int.Parse(sdecimal);                                

                decimal d10 = 1;

                for (int k = 0; k < sdecimal.Length; k++)
                {
                    d10 = d10 * 10;
                }

                decimal ddecimal_part = (decimal)decimal_part;

                decimal ddec=ddecimal_part/d10;

                decimal dintegral_part = (decimal)integral_part;

                decimal d = dintegral_part + ddec;

                if (ihask)
                {
                    d = d * 1000;
                }
                

                return d;
            }
        }

        public static string EscapeLikeValue(string value)
        {
            StringBuilder sb = new StringBuilder(value.Length);
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                switch (c)
                {
                    case ']':
                    case '[':
                    case '%':
                    case '*':
                        sb.Append("[").Append(c).Append("]");
                        break;
                    case '\'':
                        sb.Append("''");
                        break;
                    default:
                        sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }       
    }

    public class GetTitleAndNumberOfPagesResult
    {
        public string Title = "";
        public int NumberOfPages = -1;
    }

    public class InitializedBackgroundWorker : System.ComponentModel.BackgroundWorker
    {
        public InitializedBackgroundWorker()
            : base()
        {
        }

        protected override void OnDoWork(System.ComponentModel.DoWorkEventArgs e)
        {
            //MainHelper.ActionStopped = false;

            base.OnDoWork(e);
        }
    }   
}

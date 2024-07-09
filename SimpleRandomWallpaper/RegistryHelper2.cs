using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace SimpleRandomWallpaper
{
    class RegistryHelper2
    {
        public static string GetKeyValue(string AppRegKey, string keystr)
        {
            try
            {
                RegistryKey key = Registry.CurrentUser;

                key = key.OpenSubKey("Software",true);

                RegistryKey key2 = null;

                try
                {
                    key2 = key.OpenSubKey("softpcapps Software",true);

                    if (key2 == null)
                    {
                        key2 = key.CreateSubKey("softpcapps Software");
                    }

                    RegistryKey key3 = key2.OpenSubKey(AppRegKey,true);

                    try
                    {

                        if (key3 == null)
                        {
                            key3 = key2.CreateSubKey(AppRegKey);
                        }

                        if (key3 != null)
                        {
                            return key3.GetValue(keystr, "").ToString();
                        }

                        return "";

                    }
                    finally
                    {
                        if (key3 != null)
                        {
                            key3.Close();
                        }
                    }

                }
                finally
                {
                    if (key2 != null)
                    {
                        key2.Close();
                    }
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        public static bool SetKeyValue(string AppRegKey, string keystr, string keyvalue)
        {
            RegistryKey key = Registry.CurrentUser;

            key = key.OpenSubKey("Software");

            RegistryKey key2 = null;

            try
            {
                key2 = key.OpenSubKey("softpcapps Software",true);

                if (key2 == null)
                {
                    key2 = key.CreateSubKey("softpcapps Software");
                }

                RegistryKey key3 = key2.OpenSubKey(AppRegKey, true);

                try
                {

                    if (key3 == null)
                    {
                        key3 = key2.CreateSubKey(AppRegKey);

                        key3.Close();
                    }

                    key3 = key2.OpenSubKey(AppRegKey, true);

                    if (key3 != null)
                    {
                        key3.SetValue(keystr, keyvalue);
                    }

                }
                finally
                {
                    if (key3 != null)
                    {
                        key3.Close();
                    }
                }
            }
            finally
            {
                if (key2 != null)
                {
                    key2.Close();
                }
            }

            return true;
        }

        public static string GetKeyValueLMLowPriv(string AppRegKey, string keystr)
        {
            try
            {
                RegistryKey key = Registry.LocalMachine;

                key = key.OpenSubKey("Software");

                RegistryKey key2 = null;

                try
                {
                    key2 = key.OpenSubKey("softpcapps Software");

                    RegistryKey key3 = key2.OpenSubKey(AppRegKey);

                    try
                    {

                        if (key3 != null)
                        {
                            return key3.GetValue(keystr, "").ToString();
                        }

                        return "";

                    }
                    finally
                    {
                        if (key3 != null)
                        {
                            key3.Close();
                        }
                    }

                }
                finally
                {
                    if (key2 != null)
                    {
                        key2.Close();
                    }
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        public static bool RegisterKeyLmFolder(string app, string key, List<string> lstval)
        {
            try
            {
                if (app.StartsWith("\"") || app.StartsWith("'"))
                {
                    app = app.Substring(1, app.Length - 2);
                }

                if (key.StartsWith("\"") || key.StartsWith("'"))
                {
                    key = key.Substring(1, key.Length - 2);
                }

                RegistryKey keym = Registry.LocalMachine;
                RegistryKey keym2 = Registry.LocalMachine;

                keym = keym.OpenSubKey("Software\\softpcapps Software", true);

                if (keym == null)
                {
                    keym = Registry.LocalMachine.CreateSubKey("SOFTWARE\\softpcapps Software");
                }

                keym2 = keym.OpenSubKey(app, true);

                if (keym2 == null)
                {
                    keym2 = keym.CreateSubKey(app);
                }

                keym = keym2.OpenSubKey(key, true);

                if (keym == null)
                {
                    keym = keym2.CreateSubKey(key);
                }

                string[] valnames = keym.GetValueNames();

                for (int k = 0; k < valnames.Length; k++)
                {
                    keym.DeleteValue(valnames[k]);
                }

                for (int k = 0; k < lstval.Count; k++)
                {
                    keym.SetValue("#" + k.ToString(), lstval[k]);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool RegisterKeyFolder(string app, string key, List<string> lstval)
        {
            try
            {
                if (app.StartsWith("\"") || app.StartsWith("'"))
                {
                    app = app.Substring(1, app.Length - 2);
                }

                if (key.StartsWith("\"") || key.StartsWith("'"))
                {
                    key = key.Substring(1, key.Length - 2);
                }

                RegistryKey keym = Registry.CurrentUser;
                RegistryKey keym2 = Registry.CurrentUser;

                keym = keym.OpenSubKey("Software\\softpcapps Software", true);

                if (keym == null)
                {
                    keym = Registry.CurrentUser.CreateSubKey("SOFTWARE\\softpcapps Software");
                }

                keym2 = keym.OpenSubKey(app, true);

                if (keym2 == null)
                {
                    keym2 = keym.CreateSubKey(app);
                }

                keym = keym2.OpenSubKey(key, true);

                if (keym == null)
                {
                    keym = keym2.CreateSubKey(key);
                }

                string[] valnames = keym.GetValueNames();

                for (int k = 0; k < valnames.Length; k++)
                {
                    keym.DeleteValue(valnames[k]);
                }

                for (int k = 0; k < lstval.Count; k++)
                {
                    keym.SetValue("#" + k.ToString(), lstval[k]);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}

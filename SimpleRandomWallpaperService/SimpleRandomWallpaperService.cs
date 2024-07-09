using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;

namespace SimpleRandomWallpaperService
{
    public partial class SimpleRandomWallpaperService : ServiceBase
    {
        public System.Timers.Timer timFirstChange = new System.Timers.Timer();
        public System.Timers.Timer timChange = new System.Timers.Timer();

        public SimpleRandomWallpaperService()
        {
            InitializeComponent();

            timFirstChange.Elapsed += TimFirstChange_Elapsed;
            timChange.Elapsed += TimChange_Elapsed;
        }

        private void TimChange_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //Module.DebugWrite("tim change elapsed");

            try
            {
                RunChangeWallpaper();
            }
            catch { }
        }

        private void TimFirstChange_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {                       
            
            try
            {
                RunChangeWallpaper();
            }
            catch { }

            timChange.Enabled = true;
            timChange.Start();
            
            timFirstChange.Enabled = false;
            timFirstChange.Stop();
        }

        protected override void OnStart(string[] args)
        {
            try
            {       
                while (!System.Diagnostics.Process.GetCurrentProcess().Responding)
                {
                    System.Windows.Forms.Application.DoEvents();
                }

                /*
                while (CPUUsageMeter.CPUUsageHigh)
                {
                    System.Threading.Thread.Sleep(3000);
                }
                */

                try
                {
                    string sid = ProcessExtensions.GetCurrentSessionUsername();

                    string account = new System.Security.Principal.SecurityIdentifier(sid).Translate(typeof(System.Security.Principal.NTAccount)).ToString();

                    string lastchange = RegistryHelper2.GetKeyValueWithSid(sid, Module.ApplicationName, "LastChange");

                    string interval = RegistryHelper2.GetKeyValueWithSid(sid, Module.ApplicationName, "Interval");

                    string enabled = RegistryHelper2.GetKeyValueWithSid(sid, Module.ApplicationName, "Enabled");

                    string enabledall = RegistryHelper2.GetKeyValueLM(Module.ApplicationName, "EnabledForAll");

                    if ((enabled == string.Empty || enabled == bool.FalseString) && enabledall != bool.TrueString)
                    {
                        return;
                    }
                    else if (enabled == bool.TrueString)
                    {

                    }
                    else if (enabledall == bool.TrueString)
                    {
                        interval = RegistryHelper2.GetKeyValueLMLowPriv(Module.ApplicationName, "Interval");
                    }
                    else
                    {
                        return;
                    }

                    TimeSpan tsday = new TimeSpan(1, 0, 0, 0);

                    int iInterval = (int)tsday.TotalSeconds;

                    try
                    {
                        iInterval = int.Parse(interval);
                    }
                    catch
                    {

                    }

                    Module.DebugWrite("interval=" + interval);

                    iInterval = iInterval * 1000;

                    if (lastchange == string.Empty)
                    {
                        Module.DebugWrite("1");

                        RunChangeWallpaper();

                        timChange.Interval = iInterval;
                        timChange.Enabled = true;
                        timChange.Start();
                    }
                    else
                    {
                        int iFirstInterval = 0;

                        try
                        {
                            DateTime dtlast = DateTime.FromFileTimeUtc(long.Parse(lastchange));

                            DateTime dtnow = DateTime.Now;

                            TimeSpan ts = dtnow - dtlast;

                            iFirstInterval = iInterval - ((int)ts.TotalSeconds * 1000);

                            if (iFirstInterval <= 0)
                            {
                                Module.DebugWrite("2");

                                RunChangeWallpaper();

                                Module.DebugWrite("2b interval=" + iInterval.ToString());

                                timChange.Interval = iInterval;
                                timChange.Enabled = true;
                                timChange.Start();
                            }
                            else
                            {
                                Module.DebugWrite("3");

                                timFirstChange.Interval = iFirstInterval;
                                timFirstChange.Enabled = true;
                                timFirstChange.Start();

                                timChange.Interval = iInterval;
                            }
                        }
                        catch (Exception ex)
                        {
                            Module.DebugWrite("4" + ex.ToString());
                        }
                    }
                }
                catch (Exception exm)
                {

                    Module.DebugWrite("5" + exm.ToString());
                }
            }
            catch (Exception ex)
            {
                //System.IO.File.WriteAllText(@"c:\1\srwservice.txt", ex.ToString());
            }
        }

        private void RunChangeWallpaper()
        {
            while (CPUUsageMeter.CPUUsageHigh)
            {
                System.Threading.Thread.Sleep(3000);
            }

            try
            {
                Module.DebugWrite("run change wallpaper");

                string sid = ProcessExtensions.GetCurrentSessionUsername();

                string enabled = RegistryHelper2.GetKeyValueWithSid(sid, Module.ApplicationName, "Enabled");

                string enabledall = RegistryHelper2.GetKeyValueLM(Module.ApplicationName, "EnabledForAll");

                bool enabledforall = false;

                if (((enabled == bool.FalseString) || (enabled == string.Empty)) && (enabledall != bool.TrueString))
                {
                    return;
                }
                else if (enabled == bool.TrueString)
                {

                }
                else if (enabledall == bool.TrueString)
                {
                    enabledforall = true;
                }

                string dir = RegistryHelper2.GetKeyValueLM(Module.ApplicationName, "InstallationDirectory");

                string exepath = System.IO.Path.Combine(dir, "SimpleRandomWallpaper.exe");

                if (enabledforall)
                {
                    sid = "allusers";
                }

                string args = " -setwallpaper \"" + sid + "\"";

                ProcessExtensions.StartProcessAsCurrentUser(exepath, args, null, false);

                /*
                System.Diagnostics.Process pr = new Process();
                pr.StartInfo.FileName = "\"" + exepath + "\"";
                pr.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                pr.StartInfo.Arguments = "-setwallpaper " + sid;
                pr.Start();
                */
            }
            catch (Exception ex){

                Module.DebugWrite("run change wallpaper ex" + ex.ToString());
            }
        }

        protected override void OnStop()
        {
        }
                
        protected override void OnSessionChange(SessionChangeDescription changeDescription)
        {
            try
            {
                Module.DebugWrite("on session change");

                base.OnSessionChange(changeDescription);

                OnStart(new string[] { });
            }
            catch { }
        }
    }
}

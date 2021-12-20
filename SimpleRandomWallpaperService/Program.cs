using System;
using System.Collections.Generic;
using System.ServiceProcess;
using System.Text;

namespace SimpleRandomWallpaperService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ExceptionHandlersHelper.AddUnhandledExceptionHandlers();

            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new SimpleRandomWallpaperService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}

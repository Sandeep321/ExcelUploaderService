using System;
using System.ServiceProcess;

namespace Excel_upload
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {


            try
            {
                if (Environment.UserInteractive)
                {
                    var service = new ExcelUploader();
                    service.KickStart(args);
                    return;
                }
                else
                {
                    ServiceBase[] ServicesToRun;
                    ServicesToRun = new ServiceBase[]
                    {
                    new ExcelUploader()
                    };
                    ServiceBase.Run(ServicesToRun);
                }

            }
            catch (Exception ex)
            {
                Logger.Write(MessageType.Fatal, "Error occured : " + ex);
            }
        }
    }
}

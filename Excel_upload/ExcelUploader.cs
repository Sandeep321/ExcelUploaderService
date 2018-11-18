using System;
using System.ServiceProcess;
using System.Configuration;
using System.Reflection;
using log4net;

namespace Excel_upload
{
    public partial class ExcelUploader : ServiceBase
    {
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        string WatchPath1 = ConfigurationManager.AppSettings["WatchPath1"];
        string WatchPath2 = ConfigurationManager.AppSettings["WatchPath2"];
        string template = ConfigurationManager.AppSettings["templatename"];
        private readonly Uploader _uploader;
        public ExcelUploader()
        {
            try
            {
                InitializeComponent();
                _uploader = new Uploader();
                fileWatcherWatchDdriveArticleimagefolder.Created += _uploader.fileWatcherWatchDdriveArticleimagefolder_Created;
                fileWatcherWatchDDriveMYdataFolder.Created += _uploader.fileWatcherWatchDDriveMYdataFolder_Created;                
            }
            catch (Exception ex)
            {
                Logger.Write(MessageType.Fatal, "Error Occured : " + ex);
            }
        }

        public void KickStart(string[] args)
        {
            OnStart(args);
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                Logger.Write(MessageType.Info, "--- STARTED ---");
                fileWatcherWatchDdriveArticleimagefolder.Path = WatchPath1;
                fileWatcherWatchDDriveMYdataFolder.Path = WatchPath2;
                _uploader.Upload();
            }
            catch (Exception ex)
            {
                Logger.Write(MessageType.Fatal, "Error Occured : " + ex);
            }
        }
        protected override void OnStop()
        {
            try
            {
                Logger.Write(MessageType.Info, "--- STOPPED ---");
            }
            catch (Exception ex)
            {
                Logger.Write(MessageType.Fatal, "Error Occured : " + ex);
            }
        }
    }
}

namespace Excel_upload
{
    partial class ExcelUploader
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code
       
        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.fileWatcherWatchDdriveArticleimagefolder = new System.IO.FileSystemWatcher();
            this.fileWatcherWatchDDriveMYdataFolder = new System.IO.FileSystemWatcher();
            ((System.ComponentModel.ISupportInitialize)(this.fileWatcherWatchDdriveArticleimagefolder)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.fileWatcherWatchDDriveMYdataFolder)).BeginInit();
            // 
            // fileWatcherWatchDdriveArticleimagefolder
            // 
            this.fileWatcherWatchDdriveArticleimagefolder.EnableRaisingEvents = true;
            // 
            // fileWatcherWatchDDriveMYdataFolder
            // 
            this.fileWatcherWatchDDriveMYdataFolder.EnableRaisingEvents = true;
            this.ServiceName = "Service1";
            ((System.ComponentModel.ISupportInitialize)(this.fileWatcherWatchDdriveArticleimagefolder)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.fileWatcherWatchDDriveMYdataFolder)).EndInit();
            components = new System.ComponentModel.Container();
        }

        #endregion
        private System.IO.FileSystemWatcher fileWatcherWatchDdriveArticleimagefolder;
        private System.IO.FileSystemWatcher fileWatcherWatchDDriveMYdataFolder;
    }
}

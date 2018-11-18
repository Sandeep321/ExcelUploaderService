using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Timers;

namespace Excel_upload
{
    public class Uploader
    {
        static string WatchPath1 = ConfigurationManager.AppSettings["WatchPath1"];
        static string WatchPath2 = ConfigurationManager.AppSettings["WatchPath2"];
        string template = ConfigurationManager.AppSettings["templatename"];
        public bool Upload()
        {

            string filename = "MYdata.xlsx";
            string excelPath = WatchPath1 + "\\" + filename;
            string conString = string.Empty;
            string databaseconString = string.Empty;
            string extension = Path.GetExtension(".xlsx");
            if (extension == "") return false;
            //switch (extension)
            //{
            //    case ".xls": //Excel 97-03
            //        conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
            //        break;
            //    case ".xlsx": //Excel 07 or higher
            //        conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
            //        break;

            //}
            databaseconString = ConfigurationManager.ConnectionStrings["Database_connect"].ConnectionString;
            conString = String.Concat("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=", excelPath.ToString(), ";Extended Properties='Excel 12.0 Xml;HDR=YES';");
            conString = string.Format(conString, excelPath);
            using (OleDbConnection excel_con = new OleDbConnection(conString))
            {
                excel_con.Open();
                string sheet1 = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                DataTable dtExcelData = new DataTable();

                //[OPTIONAL]: It is recommended as otherwise the data will be considered as String by default.
                //dtExcelData.Columns.AddRange(new DataColumn[3] { new DataColumn("Id", typeof(int)),
                //new DataColumn("Name", typeof(string)),
                //new DataColumn("Salary", typeof(decimal)) });

                using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", excel_con))
                {
                    oda.Fill(dtExcelData);
                }
                excel_con.Close();

                string consString = ConfigurationManager.ConnectionStrings["Database_connect"].ConnectionString;
                using (SqlConnection con = new SqlConnection(consString))
                {
                    con.Close();
                    con.Open();
                    var templateformat = "";
                    string query = "select TemplateFunction from Template where Name ='Test1'";
                    SqlCommand cmd = new SqlCommand();
                    cmd = new SqlCommand(query, con);
                    SqlDataReader DR1 = cmd.ExecuteReader();
                    if (DR1.Read())
                    {
                        if (DR1.GetValue(0) == null) return false;
                        templateformat = DR1.GetValue(0).ToString();
                        string[] db_mapping = templateformat.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            //Set the database table name
                            sqlBulkCopy.DestinationTableName = "dbo.Invoice_Mstr";
                            DateTime dt = new DateTime();
                            var tt = DateTime.TryParse(dtExcelData.Columns[4].ToString(), out dt);
                            //[OPTIONAL]: Map the Excel columns with that of the database table
                            sqlBulkCopy.ColumnMappings.Add(dtExcelData.Columns[1].ToString(), db_mapping[0].ToString());
                            sqlBulkCopy.ColumnMappings.Add(dtExcelData.Columns[2].ToString(), db_mapping[1].ToString());
                            sqlBulkCopy.ColumnMappings.Add(dtExcelData.Columns[3].ToString(), db_mapping[2].ToString());
                            sqlBulkCopy.ColumnMappings.Add(dtExcelData.Columns[4].ToString(), db_mapping[3].ToString());
                            sqlBulkCopy.ColumnMappings.Add(dtExcelData.Columns[5].ToString(), db_mapping[4].ToString());
                            sqlBulkCopy.ColumnMappings.Add(dtExcelData.Columns[6].ToString(), db_mapping[5].ToString());
                            sqlBulkCopy.ColumnMappings.Add(dtExcelData.Columns[7].ToString(), db_mapping[6].ToString());
                            con.Close();
                            con.Open();
                            sqlBulkCopy.WriteToServer(dtExcelData);
                            con.Close();
                        }

                    }

                }
            }
            return true;
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("Service is recall at " + DateTime.Now);
        }
        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
        public void fileWatcherWatchDDriveMYdataFolder_Created(object sender, FileSystemEventArgs e)
        {
            try
            {
                System.Threading.Thread.Sleep(70000);
                //Then we need to check file is exist or not which is created.
                if (CheckFileExistance(WatchPath2, e.Name))
                {
                    //Then write code for log detail of file in text file.
                    CreateTextFile(WatchPath2, e.Name);
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        public void fileWatcherWatchDdriveArticleimagefolder_Created(object sender, FileSystemEventArgs e)
        {


            try
            {
                System.Threading.Thread.Sleep(70000);
                //Then we need to check file is exist or not which is created.
                if (CheckFileExistance(WatchPath1, e.Name))
                {
                    //Then write code for log detail of file in text file.
                    CreateTextFile(WatchPath1, e.Name);
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }


        }


        #region Method
        private bool CheckFileExistance(string FullPath, string FileName)
        {
            // Get the subdirectories for the specified directory.'
            bool IsFileExist = false;
            DirectoryInfo dir = new DirectoryInfo(FullPath);
            if (!dir.Exists)
                IsFileExist = false;
            else
            {
                string FileFullPath = Path.Combine(FullPath, FileName);
                if (File.Exists(FileFullPath))
                    IsFileExist = true;
            }
            return IsFileExist;


        }
        private void CreateTextFile(string FullPath, string FileName)
        {
            StreamWriter SW;
            if (!File.Exists(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "txtStatus_" + DateTime.Now.ToString("yyyyMMdd") + ".txt")))
            {
                SW = File.CreateText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "txtStatus_" + DateTime.Now.ToString("yyyyMMdd") + ".txt"));
                SW.Close();
            }
            using (SW = File.AppendText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "txtStatus_" + DateTime.Now.ToString("yyyyMMdd") + ".txt")))
            {
                SW.WriteLine("File Created with Name: " + FileName + " at this location: " + FullPath);
                SW.Close();
            }
        }
        public static void Create_ServiceStoptextfile()
        {
            string Destination = "E:\\Projects\\BulK_upload_services\\Excel_upload\\articleimg\\FileWatcherWinService";
            StreamWriter SW;
            if (Directory.Exists(Destination))
            {
                Destination = System.IO.Path.Combine(Destination, "txtServiceStop_" + DateTime.Now.ToString("yyyyMMdd") + ".txt");
                if (!File.Exists(Destination))
                {
                    SW = File.CreateText(Destination);
                    SW.Close();
                }
            }
            using (SW = File.AppendText(Destination))
            {
                SW.Write("\r\n\n");
                SW.WriteLine("Service Stopped at: " + DateTime.Now.ToString("dd-MM-yyyy H:mm:ss"));
                SW.Close();
            }
        }
        #endregion Method
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace IFTGxmlToFNC2
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        #region settings
        public static string FNC2ConnectionString = @"Data Source=CGM-APP11\SQLCGMAPP11;Initial Catalog=FNC2;";
        public static bool ServiceIsActive;            // флаг для запуска и остановки потока

        public static string user = "PSMExchangeUser"; // логин для базы обмена файлами 
        public static string password = "PSM_123456";  // пароль для базы обмена файлами 

        static object ServiceLogLocker = new object();

        #endregion

        // Лог сервиса
        static void ServiceLog(string Message)
        {
            lock (ServiceLogLocker)
            {
                try
                {
                    string path = AppDomain.CurrentDomain.BaseDirectory + "\\Log\\Service";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    string filename = path + "\\ServiceThread_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
                    if (!System.IO.File.Exists(filename))
                    {
                        using (StreamWriter sw = System.IO.File.CreateText(filename))
                        {
                            sw.WriteLine(DateTime.Now + ": " + Message);
                        }
                    }
                    else
                    {
                        using (StreamWriter sw = System.IO.File.AppendText(filename))
                        {
                            sw.WriteLine(DateTime.Now + ": " + Message);
                        }
                    }
                }
                catch
                {

                }
            }
        }

        public static void WaitingIFTGXML()
        {
            while (ServiceIsActive)
            {
                try
                {
                    string IFTGFolder = ConfigurationManager.AppSettings["FolderIn"];
                    // поолучаем список всех файлов в папке
                    string[] Files = Directory.GetFiles(IFTGFolder, "*.xml");

                    foreach (string file in Files)
                    {
                        ServiceLog($"File: {file}");
                        MoveXMLToFNC2(file);
                    }
                }
                catch(Exception ex)
                {
                    ServiceLog(ex.ToString());
                }
                Thread.Sleep(1000); 
            }
        }

        public static void MoveXMLToFNC2(string file)
        {
            //string fileDatestr = System.IO.File.GetCreationTime(file).ToString(); // дата создания файла
            string fileDatestr = System.IO.File.GetCreationTime(file).ToString(); // дата создания файла
           
            /*
            DateTime dateValue;

            string format = "YYYY-MM-DD hh:mm:tt";

            if (DateTime.TryParse(fileDatestr, out dateValue))
                ServiceLog($"Converted '{fileDatestr}' to {dateValue}");
            else
                ServiceLog($"Unable to parse '{fileDatestr}'.");
                */

            string fileFolder = ConfigurationManager.AppSettings["FolderIn"];
            // обрезаем только имя текущего файла
            string FileName = file.Substring(fileFolder.Length + 1);

            string FNC2ConnectionString = ConfigurationManager.ConnectionStrings["FNC2Connection"].ConnectionString;
            FNC2ConnectionString = String.Concat(FNC2ConnectionString, $"User Id = {user}; Password = {password}");

            try
            {
                using (SqlConnection FNC2Connection = new SqlConnection(FNC2ConnectionString))
                {
                    FNC2Connection.Open();

                    SqlCommand SqlInsertCommand = FNC2Connection.CreateCommand();
                    SqlInsertCommand.CommandText = "INSERT INTO FileExchangeOMLDB VALUES (@filename, @file, @filedate)";
                    //SqlInsertCommand.CommandText = "INSERT INTO FileExchangeOMLDB VALUES (@filename, @file, GETDATE())";

                    // создаем параметры для инсерта
                    SqlParameter filenameParam = new SqlParameter("@filename", SqlDbType.NVarChar, 100);

                    //SqlParameter filedateParam = new SqlParameter("@filedate", SqlDbType.NVarChar, 50);
                    SqlParameter filedateParam = new SqlParameter("@filedate", SqlDbType.DateTime);


                    // добавляем параметры к команде
                    SqlInsertCommand.Parameters.Add(filenameParam);

                    SqlInsertCommand.Parameters.Add(filedateParam);

                    // массив для хранения бинарных данных файла
                    byte[] fileData;
                    using(FileStream fs =new FileStream(file, FileMode.Open))
                    {
                        fileData = new byte[fs.Length];
                        fs.Read(fileData, 0, fileData.Length);
                        SqlParameter fileParam = new SqlParameter("@file", SqlDbType.VarBinary, Convert.ToInt32(fs.Length));
                        SqlInsertCommand.Parameters.Add(fileParam);
                    }

                    // передаем данные в команду через параметры
                    SqlInsertCommand.Parameters["@filename"].Value = FileName;
                    SqlInsertCommand.Parameters["@file"].Value = fileData;
                    SqlInsertCommand.Parameters["@filedate"].Value = fileDatestr;
                    //SqlInsertCommand.Parameters["@filedate"].Value = dateValue;

                    SqlInsertCommand.ExecuteNonQuery();
                }

                File.Delete(file);
                ServiceLog($"File {file} has been moved to FNC2.db, FileExchangeOMLDB");
                ServiceLog("");
            }
            catch(Exception ex)
            {
                ServiceLog(ex.Message);
            }

        }

        protected override void OnStart(string[] args)
        {
            ServiceIsActive = true;

            // поток, который мониторит папку и ждет xml
            Thread WaitingIFTGXMLThread = new Thread(new ThreadStart(WaitingIFTGXML));
            WaitingIFTGXMLThread.Name = " WaitingIFTGXML";
            WaitingIFTGXMLThread.Start();
            ServiceLog("Service is started");
        }

        protected override void OnStop()
        {
            ServiceLog("Service is stopped");
            ServiceIsActive = false;
        }
    }
}

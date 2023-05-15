using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace MyKofaxService
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();  
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 60000; //number in miliseconds  
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            using (SqlConnection objConn = new SqlConnection("Server=.\\SQLEXPRESS01;Database=master;Trusted_Connection=True;"))
            {
                try
                {
                    objConn.Open();
                    SqlCommand objCommand = objConn.CreateCommand();
                    objCommand.CommandText = "INSERT INTO LFDB.dbo.BatchStatistics(BatchID,BatchName,ProcessName,StatusName,UserID,StartDateTime,EndDateTime) \n" +
                                             "SELECT t1.*  FROM \n" +
                                             "(SELECT a.[BatchID],left(c.BatchName,50) as BatchName, b.[Name] as [ProcessName], " +
                                             "CASE CAST(a.[State] AS NVARCHAR(15)) WHEN '64' THEN 'Completed' WHEN '32' THEN 'Error' WHEN '4' THEN 'In Progress' WHEN '512' THEN 'In Progress' WHEN '2' THEN 'Ready' WHEN '128' THEN 'Reserved' WHEN '8' THEN 'Suspended' ELSE CAST(a.[State] AS NVARCHAR(15)) END AS StatusName, " +
                                             "a.[UserID], a.[StartDateTime], a.[EndDateTime] FROM [10.14.0.231].[RSSystem].[dbo].[BatchStatistics] a left join[10.14.0.231].[RSSystem].[dbo].[Processes] b " +
                                             "on a.[ProcessID] = b.[ProcessID] left join [10.14.0.231].[RSSystem].[dbo].[BatchCatalog] c on a.[BatchID] = c.[ExternalBatchID]) t1 LEFT JOIN LFDB.dbo.BatchStatistics t2 ON t1.BatchID = t2.BatchID and t1.BatchName = t2.BatchName and t1.ProcessName = t2.ProcessName and t1.StatusName = t2.StatusName " +
                                             "and t1.UserID = t2.UserID and t1.StartDateTime = t2.StartDateTime and t1.EndDateTime = t2.EndDateTime " +
                                             "where t2.BatchName is null";
                    int result = objCommand.ExecuteNonQuery();
                    WriteToFile("[" + DateTime.Now + "] " + "Batch statistics successfully updated " + result.ToString() + " rows.");
                }
                catch (Exception ex)
                {
                    WriteToFile("["+DateTime.Now+"] " + ex.Message);
                }
            }
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
    }
}

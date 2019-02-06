using System;
using System.Configuration;
using System.IO;


namespace DB2Excel
{
    class FileHelper
    {
        public static string filepath = "";
        public static void WriteFile()
        {
            string dataFromDB = CsvHelper.DumpDataToexcel(DbHelper.ConnectToDB());
            string file = ConfigurationManager.AppSettings["Excelpath"].ToString() + DateTime.Now.ToString("yyyyMMddTHHmmss") + ConfigurationManager.AppSettings["FileFormat"].ToString();
            File.WriteAllText(file, dataFromDB);
            filepath = file;
        }
    }
}

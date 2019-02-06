using System;
using System.Diagnostics;

namespace DB2Excel
{
    class Program
    {

        static void Main(string[] args)
        {
            FileHelper.WriteFile();
            Process.Start(FileHelper.filepath);
            ExcelHelper.ConvertCSVToExcel();
            Environment.Exit(0);

        }
    }
}

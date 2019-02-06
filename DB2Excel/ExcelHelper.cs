using Microsoft.Office.Interop.Excel;
using System;
using System.Configuration;

namespace DB2Excel
{
    class ExcelHelper
    {
        public static void ConvertCSVToExcel()
        {
            if (ConfigurationManager.AppSettings["GenerateExcel"].ToString().Equals("True"))
            {
                string ExcelFileName = ConfigurationManager.AppSettings["Excelpath"].ToString()+"Excel" + DateTime.Now.ToString("yyyyMMddTHHmmss") + ".xlsx";
                Application app = new Application();
                Workbook wb = app.Workbooks.Open(FileHelper.filepath);
                wb.SaveAs(ExcelFileName, XlFileFormat.xlOpenXMLWorkbook, AccessMode: XlSaveAsAccessMode.xlExclusive);
                wb.Close();
                app.Quit();
            }
        }
    }
}

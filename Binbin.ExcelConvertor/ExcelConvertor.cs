using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace Binbin.ExcelConvertor
{
    public static class ExcelConvertor
    {
        public static void Convert(string fileToConvert, string saveDirectory, string saveFileName, XlFileFormat fileFormat)
        {
            string saveFilePath = saveDirectory + saveFileName + ".xls";
            var application = new Application();
            var workbook = application.Workbooks.Open(fileToConvert);
            workbook.SaveAs(saveFilePath, fileFormat);
            workbook.Close();
            application.Quit();
        }
    }
}

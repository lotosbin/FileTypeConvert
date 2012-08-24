using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Binbin.ExcelConvertor;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = Directory.GetCurrentDirectory() + @"\Template.xlsx";
            var saveDirectory = Directory.GetCurrentDirectory();
            var saveFileName = @"\export";
            var fileFormat = XlFileFormat.xlExcel8;

            ExcelConvertor.Convert(fileName, saveDirectory, saveFileName, fileFormat);
        }

       
    }
}

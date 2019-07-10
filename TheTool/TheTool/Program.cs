using OfficeOpenXml;
using System;
using System.IO;

namespace TheTool
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Product");

                var product = new Product { };

                p.SaveAs(new FileInfo(@"to_import.xlsx"));
            }
        }

        static void PrintProductHeader()
        {

        }

        static void PrintProductValues()
        {

        }

        static void PrintAttributesHeader()
        {

        }

        static void PrintAttributesValues()
        {

        }
    }
}

using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace TheTool
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var p = new ExcelPackage(new FileInfo(@"to_import.xlsx")))
            {
                var ws = p.Workbook.Worksheets["Product"];

                var products = new List<Product> {
                    new Product
                    {
                        Name = "name1",
                        Price = 1111,
                        SKU = 11,
                        Attributes = new Attributes
                        {
                            AgeFrom = 10,
                            Language = "укр",
                            MaxGameTime = 60,
                            MinGameTime = 10,
                            MinPlayers = 2,
                            MaxPlayers = 4
                        }
                    },
                    new Product
                    {
                        Name = "name2",
                        Price = 222,
                        SKU = 22,
                        Attributes = new Attributes
                        {
                            AgeFrom = 11,
                            Language = "укр",
                            MaxGameTime = 61,
                            MinGameTime = 11,
                            MinPlayers = 3,
                            MaxPlayers = 5
                        }
                    }
                };

                var productPrinter = new ProductPrinter(ws);
                productPrinter.PrintProducts(products);

                p.SaveAs(new FileInfo(@"to_import2.xlsx"));
            }
        }
    }
}

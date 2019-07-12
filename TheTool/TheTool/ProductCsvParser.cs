using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TheTool
{
    class ProductCsvParser
    {
        public IEnumerable<Product> Parse(string csvFilePath)
        {
            List<Product> res = new List<Product>();

            var lines = File.ReadAllLines(csvFilePath)
                .Select(a => a.Split(';'))
                .ToArray();

            foreach (var row in lines)
            {
                var p = new Product
                {
                    Name = ParseName(row[0]),
                    Price = parsePrice(row[1]),
                    SKU = ParseSku(row[2]),
                    Attributes = new Attributes
                    {
                        AgeFrom = ParseAgeFrom(row[3]),
                        ...
                    }
                }
            }
        } 
    }
}

using System;
using System.Collections.Generic;
using OfficeOpenXml;

namespace TheTool
{
    class ProductPrinter
    {
        private ExcelWorksheet worksheet;

        private int currentRowId = 1;

        private const int attributeOffset = 2;
        private const int attributeValueColumnIndex = 5;

        public ProductPrinter(ExcelWorksheet blankWorkSheet)
        {
            worksheet = blankWorkSheet;
        }

        public void PrintProducts(IEnumerable<Product> products)
        {
            foreach (var product in products)
            {
                PrintProduct(product);
            }
        }

        public void PrintProduct(Product product)
        {
            currentRowId++;

            worksheet.Cells[currentRowId, 5].Value = product.Name;
            worksheet.Cells[currentRowId, 72].Value = product.Price;
            worksheet.Cells[currentRowId, 17].Value = product.SKU;

            currentRowId += attributeOffset;
            PrintAttributes(product);

        }

        private void PrintAttributes(Product product)
        {
            worksheet.Cells[currentRowId, attributeValueColumnIndex].Value = product.Attributes.MinPlayers;
            currentRowId++;
            worksheet.Cells[currentRowId, attributeValueColumnIndex].Value = product.Attributes.MaxPlayers;
            currentRowId++;
            worksheet.Cells[currentRowId, attributeValueColumnIndex].Value = product.Attributes.AgeFrom;
            currentRowId++;
            worksheet.Cells[currentRowId, attributeValueColumnIndex].Value = product.Attributes.MinGameTime;
            currentRowId++;
            worksheet.Cells[currentRowId, attributeValueColumnIndex].Value = product.Attributes.MaxGameTime;
            currentRowId++;
            worksheet.Cells[currentRowId, attributeValueColumnIndex].Value = product.Attributes.Language;
        }
    }
}

using System;
using OfficeOpenXml;

namespace TheTool
{
    class ProductPrinter
    {
        private ExcelWorksheet worksheet;

        private int currentRowId = 0;

        public ProductPrinter(ExcelWorksheet blankWorkSheet)
        {
            worksheet = blankWorkSheet;
        }

        public void PrintProduct(Product product)
        {
            PrintProductHeader();
        }

        private void PrintProductHeader()
        {
            worksheet.Cells[currentRowId, 0].Value = "ProductId";
            worksheet.Cells[currentRowId, 0].Value = "ProductType";
            worksheet.Cells[currentRowId, 0].Value = "ParentGroupedProductId";
            worksheet.Cells[currentRowId, 0].Value = "VisibleIndividually";
            worksheet.Cells[currentRowId, 0].Value = "Name";
            worksheet.Cells[currentRowId, 0].Value = "ShortDescription";
            worksheet.Cells[currentRowId, 0].Value = "FullDescription";
            worksheet.Cells[currentRowId, 0].Value = "Vendor";
            worksheet.Cells[currentRowId, 0].Value = "ProductTemplate";
            worksheet.Cells[currentRowId, 0].Value = "ShowOnHomePage";
            worksheet.Cells[currentRowId, 0].Value = "MetaKeywords";
            worksheet.Cells[currentRowId, 0].Value = "MetaDescription";
            worksheet.Cells[currentRowId, 0].Value = "MetaTitle";
            worksheet.Cells[currentRowId, 0].Value = "SeName";
            worksheet.Cells[currentRowId, 0].Value = "AllowCustomerReviews";
            worksheet.Cells[currentRowId, 0].Value = "Published";
            worksheet.Cells[currentRowId, 0].Value = "SKU";
            worksheet.Cells[currentRowId, 0].Value = "ManufacturerPartNumber";
            worksheet.Cells[currentRowId, 0].Value = "Gtin";
            worksheet.Cells[currentRowId, 0].Value = "IsGiftCard";
            worksheet.Cells[currentRowId, 0].Value = "GiftCardType";
            worksheet.Cells[currentRowId, 0].Value = "OverriddenGiftCardAmount";
            worksheet.Cells[currentRowId, 0].Value = "RequireOtherProducts";
            worksheet.Cells[currentRowId, 0].Value = "RequiredProductIds";
            worksheet.Cells[currentRowId, 0].Value = "AutomaticallyAddRequiredProducts";
            worksheet.Cells[currentRowId, 0].Value = "IsDownload";
            worksheet.Cells[currentRowId, 0].Value = "UnlimitedDownloads";
            worksheet.Cells[currentRowId, 0].Value = "MaxNumberOfDownloads";
            worksheet.Cells[currentRowId, 0].Value = "DownloadActivationType";
            worksheet.Cells[currentRowId, 0].Value = "HasSampleDownload";
            worksheet.Cells[currentRowId, 0].Value = "SampleDownloadId";
            worksheet.Cells[currentRowId, 0].Value = "HasUserAgreement";
            worksheet.Cells[currentRowId, 0].Value = "UserAgreementText";
            worksheet.Cells[currentRowId, 0].Value = "IsRecurring";
            worksheet.Cells[currentRowId, 0].Value = "RecurringCycleLength";
            worksheet.Cells[currentRowId, 0].Value = "RecurringCyclePeriod";
            worksheet.Cells[currentRowId, 0].Value = "RecurringTotalCycles";
            worksheet.Cells[currentRowId, 0].Value = "IsRental";
            worksheet.Cells[currentRowId, 0].Value = "RentalPriceLength";
            worksheet.Cells[currentRowId, 0].Value = "RentalPricePeriod";
            worksheet.Cells[currentRowId, 0].Value = "IsShipEnabled";
            worksheet.Cells[currentRowId, 0].Value = "IsFreeShipping";
            worksheet.Cells[currentRowId, 0].Value = "ShipSeparately";
            worksheet.Cells[currentRowId, 0].Value = "AdditionalShippingCharge";
            worksheet.Cells[currentRowId, 0].Value = "DeliveryDate";
            worksheet.Cells[currentRowId, 0].Value = "IsTaxExempt";
            worksheet.Cells[currentRowId, 0].Value = "TaxCategory";
            worksheet.Cells[currentRowId, 0].Value = "IsTelecommunicationsOrBroadcastingOrElectronicServices";
            worksheet.Cells[currentRowId, 0].Value = "ManageInventoryMethod";
            worksheet.Cells[currentRowId, 0].Value = "ProductAvailabilityRange";
            worksheet.Cells[currentRowId, 0].Value = "UseMultipleWarehouses";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "StockQuantity";
            worksheet.Cells[currentRowId, 0].Value = "DisplayStockAvailability";
            worksheet.Cells[currentRowId, 0].Value = "DisplayStockQuantity";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 0].Value = "WarehouseId";
        }
    }
}

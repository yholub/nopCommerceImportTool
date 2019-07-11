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
            worksheet.Cells[currentRowId, 1].Value = "ProductType";
            worksheet.Cells[currentRowId, 2].Value = "ParentGroupedProductId";
            worksheet.Cells[currentRowId, 3].Value = "VisibleIndividually";
            worksheet.Cells[currentRowId, 4].Value = "Name";
            worksheet.Cells[currentRowId, 5].Value = "ShortDescription";
            worksheet.Cells[currentRowId, 6].Value = "FullDescription";
            worksheet.Cells[currentRowId, 7].Value = "Vendor";
            worksheet.Cells[currentRowId, 8].Value = "ProductTemplate";
            worksheet.Cells[currentRowId, 9].Value = "ShowOnHomePage";
            worksheet.Cells[currentRowId, 10].Value = "MetaKeywords";
            worksheet.Cells[currentRowId, 11].Value = "MetaDescription";
            worksheet.Cells[currentRowId, 12].Value = "MetaTitle";
            worksheet.Cells[currentRowId, 13].Value = "SeName";
            worksheet.Cells[currentRowId, 14].Value = "AllowCustomerReviews";
            worksheet.Cells[currentRowId, 15].Value = "Published";
            worksheet.Cells[currentRowId, 16].Value = "SKU";
            worksheet.Cells[currentRowId, 17].Value = "ManufacturerPartNumber";
            worksheet.Cells[currentRowId, 18].Value = "Gtin";
            worksheet.Cells[currentRowId, 19].Value = "IsGiftCard";
            worksheet.Cells[currentRowId, 20].Value = "GiftCardType";
            worksheet.Cells[currentRowId, 21].Value = "OverriddenGiftCardAmount";
            worksheet.Cells[currentRowId, 22].Value = "RequireOtherProducts";
            worksheet.Cells[currentRowId, 23].Value = "RequiredProductIds";
            worksheet.Cells[currentRowId, 24].Value = "AutomaticallyAddRequiredProducts";
            worksheet.Cells[currentRowId, 25].Value = "IsDownload";
            worksheet.Cells[currentRowId, 26].Value = "UnlimitedDownloads";
            worksheet.Cells[currentRowId, 27].Value = "MaxNumberOfDownloads";
            worksheet.Cells[currentRowId, 28].Value = "DownloadActivationType";
            worksheet.Cells[currentRowId, 29].Value = "HasSampleDownload";
            worksheet.Cells[currentRowId, 30].Value = "SampleDownloadId";
            worksheet.Cells[currentRowId, 31].Value = "HasUserAgreement";
            worksheet.Cells[currentRowId, 32].Value = "UserAgreementText";
            worksheet.Cells[currentRowId, 33].Value = "IsRecurring";
            worksheet.Cells[currentRowId, 34].Value = "RecurringCycleLength";
            worksheet.Cells[currentRowId, 35].Value = "RecurringCyclePeriod";
            worksheet.Cells[currentRowId, 36].Value = "RecurringTotalCycles";
            worksheet.Cells[currentRowId, 37].Value = "IsRental";
            worksheet.Cells[currentRowId, 38].Value = "RentalPriceLength";
            worksheet.Cells[currentRowId, 39].Value = "RentalPricePeriod";
            worksheet.Cells[currentRowId, 40].Value = "IsShipEnabled";
            worksheet.Cells[currentRowId, 41].Value = "IsFreeShipping";
            worksheet.Cells[currentRowId, 42].Value = "ShipSeparately";
            worksheet.Cells[currentRowId, 43].Value = "AdditionalShippingCharge";
            worksheet.Cells[currentRowId, 44].Value = "DeliveryDate";
            worksheet.Cells[currentRowId, 45].Value = "IsTaxExempt";
            worksheet.Cells[currentRowId, 46].Value = "TaxCategory";
            worksheet.Cells[currentRowId, 47].Value = "IsTelecommunicationsOrBroadcastingOrElectronicServices";
            worksheet.Cells[currentRowId, 48].Value = "ManageInventoryMethod";
            worksheet.Cells[currentRowId, 49].Value = "ProductAvailabilityRange";
            worksheet.Cells[currentRowId, 50].Value = "UseMultipleWarehouses";
            worksheet.Cells[currentRowId, 51].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 52].Value = "StockQuantity";
            worksheet.Cells[currentRowId, 53].Value = "DisplayStockAvailability";
            worksheet.Cells[currentRowId, 54].Value = "DisplayStockQuantity";
            worksheet.Cells[currentRowId, 55].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 56].Value = "StockQuantity";
            worksheet.Cells[currentRowId, 57].Value = "DisplayStockAvailability";
            worksheet.Cells[currentRowId, 58].Value = "DisplayStockQuantity";
            worksheet.Cells[currentRowId, 59].Value = "MinStockQuantity";
            worksheet.Cells[currentRowId, 60].Value = "LowStockActivity";
            worksheet.Cells[currentRowId, 61].Value = "NotifyAdminForQuantityBelow";
            worksheet.Cells[currentRowId, 62].Value = "BackorderMode";
            worksheet.Cells[currentRowId, 63].Value = "AllowBackInStockSubscriptions";
            worksheet.Cells[currentRowId, 64].Value = "OrderMinimumQuantity";
            worksheet.Cells[currentRowId, 65].Value = "OrderMaximumQuantity";
            worksheet.Cells[currentRowId, 66].Value = "AllowedQuantities";
            worksheet.Cells[currentRowId, 67].Value = "AllowAddingOnlyExistingAttributeCombinations";
            worksheet.Cells[currentRowId, 68].Value = "NotReturnable";
            worksheet.Cells[currentRowId, 69].Value = "DisableBuyButton";
            worksheet.Cells[currentRowId, 70].Value = "DisableWishlistButton";
            worksheet.Cells[currentRowId, 71].Value = "AvailableForPreOrder";
            worksheet.Cells[currentRowId, 72].Value = "PreOrderAvailabilityStartDateTimeUtc";
            worksheet.Cells[currentRowId, 73].Value = "CallForPrice";
            worksheet.Cells[currentRowId, 74].Value = "Price";
            worksheet.Cells[currentRowId, 75].Value = "OldPrice";
            worksheet.Cells[currentRowId, 76].Value = "ProductCost";
            worksheet.Cells[currentRowId, 77].Value = "CustomerEntersPrice";
            worksheet.Cells[currentRowId, 78].Value = "MinimumCustomerEnteredPrice";
            worksheet.Cells[currentRowId, 79].Value = "MaximumCustomerEnteredPrice";
            worksheet.Cells[currentRowId, 80].Value = "BasepriceEnabled";
            worksheet.Cells[currentRowId, 81].Value = "BasepriceAmount";
            worksheet.Cells[currentRowId, 82].Value = "BasepriceUnit";
            worksheet.Cells[currentRowId, 83].Value = "BasepriceBaseAmount";
            worksheet.Cells[currentRowId, 84].Value = "BasepriceBaseUnit";
            worksheet.Cells[currentRowId, 85].Value = "MarkAsNew";
            worksheet.Cells[currentRowId, 86].Value = "MarkAsNewStartDateTimeUtc";
            worksheet.Cells[currentRowId, 87].Value = "MarkAsNewEndDateTimeUtc";
            worksheet.Cells[currentRowId, 88].Value = "Weight";
            worksheet.Cells[currentRowId, 89].Value = "Length";
            worksheet.Cells[currentRowId, 90].Value = "Width";
            worksheet.Cells[currentRowId, 91].Value = "Height";
            worksheet.Cells[currentRowId, 92].Value = "Categories";
            worksheet.Cells[currentRowId, 93].Value = "Manufacturers";
            worksheet.Cells[currentRowId, 94].Value = "ProductTags";
            worksheet.Cells[currentRowId, 95].Value = "Picture1";
            worksheet.Cells[currentRowId, 96].Value = "Picture2";
            worksheet.Cells[currentRowId, 97].Value = "Picture3";
        }
    }
}

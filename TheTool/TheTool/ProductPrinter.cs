using System;
using OfficeOpenXml;

namespace TheTool
{
    class ProductPrinter
    {
        private ExcelWorksheet worksheet;

        private int currentRowId = 1;

        public ProductPrinter(ExcelWorksheet blankWorkSheet)
        {
            worksheet = blankWorkSheet;
        }

        public void PrintProduct(Product product)
        {
            PrintProductHeader();
            PrintProductValues(product);
        }

        private void PrintProductValues(Product product)
        {
            throw new NotImplementedException();
        }

        private void PrintProductHeader()
        {
            worksheet.Cells[currentRowId, 1].Value = "ProductId";
            worksheet.Cells[currentRowId, 2].Value = "ProductType";
            worksheet.Cells[currentRowId, 3].Value = "ParentGroupedProductId";
            worksheet.Cells[currentRowId, 4].Value = "VisibleIndividually";
            worksheet.Cells[currentRowId, 5].Value = "Name";
            worksheet.Cells[currentRowId, 6].Value = "ShortDescription";
            worksheet.Cells[currentRowId, 7].Value = "FullDescription";
            worksheet.Cells[currentRowId, 8].Value = "Vendor";
            worksheet.Cells[currentRowId, 9].Value = "ProductTemplate";
            worksheet.Cells[currentRowId, 10].Value = "ShowOnHomePage";
            worksheet.Cells[currentRowId, 11].Value = "MetaKeywords";
            worksheet.Cells[currentRowId, 12].Value = "MetaDescription";
            worksheet.Cells[currentRowId, 13].Value = "MetaTitle";
            worksheet.Cells[currentRowId, 14].Value = "SeName";
            worksheet.Cells[currentRowId, 15].Value = "AllowCustomerReviews";
            worksheet.Cells[currentRowId, 16].Value = "Published";
            worksheet.Cells[currentRowId, 17].Value = "SKU";
            worksheet.Cells[currentRowId, 18].Value = "ManufacturerPartNumber";
            worksheet.Cells[currentRowId, 19].Value = "Gtin";
            worksheet.Cells[currentRowId, 20].Value = "IsGiftCard";
            worksheet.Cells[currentRowId, 21].Value = "GiftCardType";
            worksheet.Cells[currentRowId, 22].Value = "OverriddenGiftCardAmount";
            worksheet.Cells[currentRowId, 23].Value = "RequireOtherProducts";
            worksheet.Cells[currentRowId, 24].Value = "RequiredProductIds";
            worksheet.Cells[currentRowId, 25].Value = "AutomaticallyAddRequiredProducts";
            worksheet.Cells[currentRowId, 26].Value = "IsDownload";
            worksheet.Cells[currentRowId, 27].Value = "UnlimitedDownloads";
            worksheet.Cells[currentRowId, 28].Value = "MaxNumberOfDownloads";
            worksheet.Cells[currentRowId, 29].Value = "DownloadActivationType";
            worksheet.Cells[currentRowId, 30].Value = "HasSampleDownload";
            worksheet.Cells[currentRowId, 31].Value = "SampleDownloadId";
            worksheet.Cells[currentRowId, 32].Value = "HasUserAgreement";
            worksheet.Cells[currentRowId, 33].Value = "UserAgreementText";
            worksheet.Cells[currentRowId, 34].Value = "IsRecurring";
            worksheet.Cells[currentRowId, 35].Value = "RecurringCycleLength";
            worksheet.Cells[currentRowId, 36].Value = "RecurringCyclePeriod";
            worksheet.Cells[currentRowId, 37].Value = "RecurringTotalCycles";
            worksheet.Cells[currentRowId, 38].Value = "IsRental";
            worksheet.Cells[currentRowId, 39].Value = "RentalPriceLength";
            worksheet.Cells[currentRowId, 40].Value = "RentalPricePeriod";
            worksheet.Cells[currentRowId, 41].Value = "IsShipEnabled";
            worksheet.Cells[currentRowId, 42].Value = "IsFreeShipping";
            worksheet.Cells[currentRowId, 43].Value = "ShipSeparately";
            worksheet.Cells[currentRowId, 44].Value = "AdditionalShippingCharge";
            worksheet.Cells[currentRowId, 45].Value = "DeliveryDate";
            worksheet.Cells[currentRowId, 46].Value = "IsTaxExempt";
            worksheet.Cells[currentRowId, 47].Value = "TaxCategory";
            worksheet.Cells[currentRowId, 48].Value = "IsTelecommunicationsOrBroadcastingOrElectronicServices";
            worksheet.Cells[currentRowId, 49].Value = "ManageInventoryMethod";
            worksheet.Cells[currentRowId, 50].Value = "ProductAvailabilityRange";
            worksheet.Cells[currentRowId, 51].Value = "UseMultipleWarehouses";
            worksheet.Cells[currentRowId, 52].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 53].Value = "StockQuantity";
            worksheet.Cells[currentRowId, 54].Value = "DisplayStockAvailability";
            worksheet.Cells[currentRowId, 55].Value = "DisplayStockQuantity";
            worksheet.Cells[currentRowId, 56].Value = "WarehouseId";
            worksheet.Cells[currentRowId, 57].Value = "StockQuantity";
            worksheet.Cells[currentRowId, 58].Value = "DisplayStockAvailability";
            worksheet.Cells[currentRowId, 59].Value = "DisplayStockQuantity";
            worksheet.Cells[currentRowId, 60].Value = "MinStockQuantity";
            worksheet.Cells[currentRowId, 61].Value = "LowStockActivity";
            worksheet.Cells[currentRowId, 62].Value = "NotifyAdminForQuantityBelow";
            worksheet.Cells[currentRowId, 63].Value = "BackorderMode";
            worksheet.Cells[currentRowId, 64].Value = "AllowBackInStockSubscriptions";
            worksheet.Cells[currentRowId, 65].Value = "OrderMinimumQuantity";
            worksheet.Cells[currentRowId, 66].Value = "OrderMaximumQuantity";
            worksheet.Cells[currentRowId, 67].Value = "AllowedQuantities";
            worksheet.Cells[currentRowId, 68].Value = "AllowAddingOnlyExistingAttributeCombinations";
            worksheet.Cells[currentRowId, 69].Value = "NotReturnable";
            worksheet.Cells[currentRowId, 70].Value = "DisableBuyButton";
            worksheet.Cells[currentRowId, 71].Value = "DisableWishlistButton";
            worksheet.Cells[currentRowId, 72].Value = "AvailableForPreOrder";
            worksheet.Cells[currentRowId, 73].Value = "PreOrderAvailabilityStartDateTimeUtc";
            worksheet.Cells[currentRowId, 74].Value = "CallForPrice";
            worksheet.Cells[currentRowId, 75].Value = "Price";
            worksheet.Cells[currentRowId, 76].Value = "OldPrice";
            worksheet.Cells[currentRowId, 77].Value = "ProductCost";
            worksheet.Cells[currentRowId, 78].Value = "CustomerEntersPrice";
            worksheet.Cells[currentRowId, 79].Value = "MinimumCustomerEnteredPrice";
            worksheet.Cells[currentRowId, 80].Value = "MaximumCustomerEnteredPrice";
            worksheet.Cells[currentRowId, 81].Value = "BasepriceEnabled";
            worksheet.Cells[currentRowId, 82].Value = "BasepriceAmount";
            worksheet.Cells[currentRowId, 83].Value = "BasepriceUnit";
            worksheet.Cells[currentRowId, 84].Value = "BasepriceBaseAmount";
            worksheet.Cells[currentRowId, 85].Value = "BasepriceBaseUnit";
            worksheet.Cells[currentRowId, 86].Value = "MarkAsNew";
            worksheet.Cells[currentRowId, 87].Value = "MarkAsNewStartDateTimeUtc";
            worksheet.Cells[currentRowId, 88].Value = "MarkAsNewEndDateTimeUtc";
            worksheet.Cells[currentRowId, 89].Value = "Weight";
            worksheet.Cells[currentRowId, 90].Value = "Length";
            worksheet.Cells[currentRowId, 91].Value = "Width";
            worksheet.Cells[currentRowId, 92].Value = "Height";
            worksheet.Cells[currentRowId, 93].Value = "Categories";
            worksheet.Cells[currentRowId, 94].Value = "Manufacturers";
            worksheet.Cells[currentRowId, 95].Value = "ProductTags";
            worksheet.Cells[currentRowId, 96].Value = "Picture1";
            worksheet.Cells[currentRowId, 97].Value = "Picture2";
            worksheet.Cells[currentRowId, 97].Value = "Picture3";
        }
    }
}
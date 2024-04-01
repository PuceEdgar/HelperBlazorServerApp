using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FileProcessingLibrary.Services;

public class ExcelService : IExcelService
{
    public  List<string> ReadPartDataFromStream(Stream stream, int sheetNumber, int startingRow, int columnNumber, bool isMaster = false)
    {
        var wb = new XSSFWorkbook(stream);
        var sheet = wb.GetSheetAt(sheetNumber);

        return ReadColumnFromSheet(sheet, startingRow, columnNumber, isMaster);
    }

    public async Task<List<StockDetails>> ExtractStockDetailsFromExcelStream(Stream stream, FileSource fileSource)
    {
        var stockDetailsList = new List<StockDetails>();
        using var ms = new MemoryStream();
        await stream.CopyToAsync(ms);
        ms.Seek(0, SeekOrigin.Begin);

        var wb = new XSSFWorkbook(ms);

        if (fileSource == FileSource.Mfg)
        {
            GetStockDetails(stockDetailsList, wb, 0, 4, 15, 10, 19);
        }
        else
        {
            GetStockDetails(stockDetailsList, wb, 2, 8, 4, 10, 12);
        }


        return stockDetailsList;
    }

    private static void GetStockDetails(List<StockDetails> stockDetailsList, XSSFWorkbook wb, int sheetNumber, int partNoCol, int customerNumberCol, int qtyCol, int unitPriceCol)
    {
        var sheet = wb.GetSheetAt(sheetNumber);
        var enumerator = sheet.GetEnumerator();
        while (enumerator.MoveNext())
        {
            if (enumerator.Current is not IRow row
                || row.Cells.TrueForAll(c => c.CellType == CellType.Blank)
                || row.RowNum < 1
                )
            {
                continue;
            }

            stockDetailsList.Add(new StockDetails
            {
                PartNumber = row.GetCell(partNoCol)?.StringCellValue.Trim(),
                CustomerNumber = row.GetCell(customerNumberCol)?.StringCellValue.Trim(),
                Qty = (int)row.GetCell(qtyCol).NumericCellValue,
                UnitPrice = (decimal)row.GetCell(unitPriceCol).NumericCellValue
            });
        }
    }

    public IWorkbook CompareItems(GroupedStockItem groupedMFGItems, GroupedStockItem groupedRBItems)
    {
        var missingItems = new List<MissingItem>();

        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Compare Stock");

        var row = sheet.CreateRow(0);
        row.CreateCell(0).SetCellValue("RB Part");
        row.CreateCell(1).SetCellValue("RB Customer");
        row.CreateCell(2).SetCellValue("RB Qty");
        row.CreateCell(3).SetCellValue("RB Price");
        row.CreateCell(4).SetCellValue("RB Rounded Price");
        row.CreateCell(5).SetCellValue("MFG Part");
        row.CreateCell(6).SetCellValue("MFG Customer");
        row.CreateCell(7).SetCellValue("MFG Qty");
        row.CreateCell(8).SetCellValue("MFG Price");
        row.CreateCell(9).SetCellValue("QTY MATCH");
        row.CreateCell(10).SetCellValue("Calculated PRICE MATCH");
        row.CreateCell(11).SetCellValue("Rounded PRICE MATCH");
        var rowNumber = 1;


        foreach (var mfgItem in groupedMFGItems.GroupedStockList)
        {
            var rbItem = groupedRBItems.GroupedStockList.Find(i => i.CustomerNumber == mfgItem.CustomerNumber && i.PartNumber == mfgItem.PartNumber);
            if (rbItem is null)
            {
                Console.WriteLine($"Item: {mfgItem.CustomerNumber} {mfgItem.PartNumber} Exists in MFG but not in RB.");
                missingItems.Add(new MissingItem
                {
                    CustomerNumber = mfgItem.CustomerNumber!,
                    PartNumber = mfgItem.PartNumber!,
                    Qty = mfgItem.Qty,
                    MissingIn = FileSource.RbInventory,
                    ExistsIn = FileSource.Mfg
                });
                continue;
            }

            PopulateRow(sheet, rowNumber, mfgItem, rbItem);
            rowNumber++;
        }

        foreach (var rbItem in groupedRBItems.GroupedStockList)
        {
            var mfgItem = groupedMFGItems.GroupedStockList.Find(i => i.CustomerNumber == rbItem.CustomerNumber && i.PartNumber == rbItem.PartNumber);
            if (mfgItem is null)
            {
                Console.WriteLine($"Item: {rbItem.CustomerNumber} {rbItem.PartNumber} Exists in RB but not in MFG.");
                missingItems.Add(new MissingItem
                {
                    CustomerNumber = rbItem.CustomerNumber!,
                    PartNumber = rbItem.PartNumber!,
                    Qty = rbItem.Qty,
                    MissingIn = FileSource.Mfg,
                    ExistsIn = FileSource.RbInventory
                });
                continue;
            }
            PopulateRow(sheet, rowNumber, mfgItem, rbItem);
            rowNumber++;
        }


        foreach (var item in missingItems)
        {
            var itemRow = sheet.CreateRow(rowNumber);

            itemRow.CreateCell(0).SetCellValue($"Missing item! Exists in: {item.ExistsIn} but missing in: {item.MissingIn}. Part number: {item.PartNumber} for customer: {item.CustomerNumber} item count: {item.Qty}");
            rowNumber++;
        }

        return workbook;
    }

    private static void PopulateRow(ISheet sheet, int rowNumber, StockDetails mfgItem, StockDetails rbItem)
    {
        var roundedRbPrice = GetRoundedPrice(mfgItem.UnitPrice, rbItem.UnitPrice);

        var itemRow = sheet.CreateRow(rowNumber);

        itemRow.CreateCell(0).SetCellValue(rbItem.PartNumber);
        itemRow.CreateCell(1).SetCellValue(rbItem.CustomerNumber);
        itemRow.CreateCell(2).SetCellValue(rbItem.Qty);
        itemRow.CreateCell(3).SetCellValue((double)rbItem.UnitPrice);
        itemRow.CreateCell(4).SetCellValue((double)roundedRbPrice);
        itemRow.CreateCell(5).SetCellValue(mfgItem.PartNumber);
        itemRow.CreateCell(6).SetCellValue(mfgItem.CustomerNumber);
        itemRow.CreateCell(7).SetCellValue(mfgItem.Qty);
        itemRow.CreateCell(8).SetCellValue((double)mfgItem.UnitPrice);

        var qtyMatch = rbItem.Qty == mfgItem.Qty;
        itemRow.CreateCell(9).SetCellValue(qtyMatch.ToString());

        var priceMatch = rbItem.UnitPrice == mfgItem.UnitPrice;
        itemRow.CreateCell(10).SetCellValue(priceMatch.ToString());

        var roundedPriceMatch = roundedRbPrice == mfgItem.UnitPrice;
        itemRow.CreateCell(11).SetCellValue(roundedPriceMatch.ToString());
    }

    private static decimal GetRoundedPrice(decimal mfgPrice, decimal rbPrice)
    {
        int decimalCount = BitConverter.GetBytes(decimal.GetBits(mfgPrice)[3])[2];
        var rbRoundedPRice = decimal.Round(rbPrice, decimalCount, MidpointRounding.ToPositiveInfinity);
        return rbRoundedPRice;
    }

    private static List<string> ReadColumnFromSheet(ISheet sheet, int startingRow, int columnNumber, bool isMasterFile)
    {
        List<string> parts = [];

        var enumerator = sheet.GetEnumerator();

        while (enumerator.MoveNext())
        {
            if (enumerator.Current is not IRow row
                || row.Cells.TrueForAll(c => c.CellType == CellType.Blank)
                || row.RowNum < startingRow
                || (isMasterFile && row.GetCell(10).StringCellValue.Trim().StartsWith('0')))
            {
                continue;
            }

            var actionRequiredBillUpDownColumn = row.GetCell(columnNumber)?.StringCellValue.Trim();

            if (!string.IsNullOrWhiteSpace(actionRequiredBillUpDownColumn))
            {
                parts.Add(actionRequiredBillUpDownColumn);
            }
        }

        return parts.Distinct().ToList();
    }
}

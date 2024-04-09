using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FileProcessingLibrary.Services;

public class ExcelService : IExcelService
{
    public  List<string> ReadPartDataFromStream(Stream stream, int sheetNumber, int startingRow, int columnNumber, bool isMaster = false)
    {
        XSSFWorkbook wb = new(stream);
        var sheet = wb.GetSheetAt(sheetNumber);

        return ReadColumnFromSheet(sheet, startingRow, columnNumber, isMaster);
    }

    public async Task<List<StockDetails>> ExtractStockDetailsFromExcelStream(Stream stream, FileSource fileSource)
    {
        List<StockDetails> stockDetailsList = [];
        using MemoryStream ms = new();
        await stream.CopyToAsync(ms);
        ms.Seek(0, SeekOrigin.Begin);

        XSSFWorkbook wb = new(ms);

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

            stockDetailsList.Add(new()
            {
                PartNumber = row.GetCell(partNoCol)?.StringCellValue.Trim(),
                CustomerNumber = row.GetCell(customerNumberCol)?.StringCellValue.Trim(),
                Qty = (int)row.GetCell(qtyCol).NumericCellValue,
                UnitPrice = (decimal)row.GetCell(unitPriceCol).NumericCellValue,
            });
        }
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

using NPOI.SS.UserModel;

namespace FileProcessingLibrary.Services;

public interface IExcelService
{
    List<string> ReadPartDataFromStream(Stream stream, int sheetNumber, int startingRow, int columnNumber, bool isMaster = false);
    Task<List<StockDetails>> ExtractStockDetailsFromExcelStream(Stream stream, FileSource fileSource);
    IWorkbook CompareItems(GroupedStockItem groupedMFGItems, GroupedStockItem groupedRBItems);
}

using NPOI.SS.UserModel;
using System.IO.Compression;

namespace FileProcessingLibrary.Services;

public interface IExcelService
{
    Dictionary<FileSource, List<string>> ExtractPartsData(Dictionary<FileSource, ZipArchiveEntry> data);
    List<string> ReadPartDataFromStream(Stream stream, int sheetNumber, int startingRow, int columnNumber, bool isMaster = false);
    Task<List<StockDetails>> ExtractStockDetailsFromExcelStream(Stream stream, FileSource fileSource);
    Task<List<StockDetails>> ExtractStockDetailsFromCsvStream(Stream stream);
    IWorkbook CompareItems(GroupedStockItem groupedMFGItems, GroupedStockItem groupedRBItems);
}

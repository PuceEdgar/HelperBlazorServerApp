using NPOI.SS.UserModel;

namespace FileProcessingLibrary.Services;

public interface IDataComparerService
{
    List<PartExistsData> CompareData(Dictionary<FileSource, List<string>> dataDict);
    IWorkbook CompareItems(GroupedStockItem groupedMFGItems, GroupedStockItem groupedRBItems);
}

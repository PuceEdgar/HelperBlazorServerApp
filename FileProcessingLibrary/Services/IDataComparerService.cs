namespace FileProcessingLibrary.Services;

public interface IDataComparerService
{
    List<PartExistsData> CompareData(Dictionary<FileSource, List<string>> dataDict);
}

namespace FileProcessingLibrary.Services;

public class DataComparerService : IDataComparerService
{
    public List<PartExistsData> CompareData(Dictionary<FileSource, List<string>> dataDict)
    {
        var existingItems = new List<PartExistsData>();

        foreach (var part in dataDict[FileSource.Master])
        {
            CheckIfPartExists(dataDict, existingItems, part);
        }

        return existingItems;
    }

    private static void CheckIfPartExists(Dictionary<FileSource, List<string>> dataDict, List<PartExistsData> existingItems, string part)
    {
        if (dataDict[FileSource.RbBilling].Contains(part))
        {
            existingItems.Add(new PartExistsData
            {
                SourceData = "RB Billing",
                PartNumber = part,
            });
        }

        if (dataDict[FileSource.RbBacklog].Contains(part))
        {
            existingItems.Add(new PartExistsData
            {
                SourceData = "RB Backlog",
                PartNumber = part,
            });
        }

        if (dataDict[FileSource.Mfg].Contains(part))
        {
            existingItems.Add(new PartExistsData
            {
                SourceData = "MFG",
                PartNumber = part,
            });
        }
    }
}

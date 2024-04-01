using ExcelComparerLibrary;

namespace ActionNeededAndCompareBacklogApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            
            var currentDir = Directory.GetCurrentDirectory();
            var dataDictionary = ZipFileReader.ProcessZipFile($"{currentDir}/PartsZipFile.zip");
            //var filesToProcess = Directory.GetFiles(currentDir);
            //if (filesToProcess.Length == 0)
            //{
            //    return;
            //}
            //var mainFilePath = Array.Find(filesToProcess, f => f.Contains("master", StringComparison.CurrentCultureIgnoreCase));
            //var rbBacklogFile = Array.Find(filesToProcess, f => f.Contains("rb", StringComparison.CurrentCultureIgnoreCase) && f.Contains("backlog", StringComparison.CurrentCultureIgnoreCase));
            //var rbBillingFile = Array.Find(filesToProcess, f => f.Contains("rb", StringComparison.CurrentCultureIgnoreCase) && f.Contains("billing", StringComparison.CurrentCultureIgnoreCase));
            //var mfgFile = Array.Find(filesToProcess, f => f.Contains("mfg", StringComparison.CurrentCultureIgnoreCase));

            //if (string.IsNullOrWhiteSpace(mainFilePath))
            //{
            //    return;
            //}

            //var mainFileParts = ExcelReader.ReadPartNoDataFromExcel(mainFilePath, 1, 4, 6, true);
            //var rbBacklogParts = string.IsNullOrEmpty(rbBacklogFile) ? [] : ExcelReader.ReadPartNoDataFromExcel(rbBacklogFile, 0, 1, 21);
            //var rbBillingParts = string.IsNullOrEmpty(rbBillingFile) ? [] : ExcelReader.ReadPartNoDataFromExcel(rbBillingFile, 0, 4, 21);
            //var mfgParts = string.IsNullOrEmpty(mfgFile) ? [] : ExcelReader.ReadPartNoDataFromExcel(mfgFile, 0, 1, 4);

            var existingParts = DataComparer.CheckIfPartExists(dataDictionary);
            Console.WriteLine("Done");
            Console.WriteLine($"Existing part count: {existingParts.Count}");
            Console.ReadLine();
        }
    }
}

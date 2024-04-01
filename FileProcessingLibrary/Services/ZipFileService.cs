using System.IO.Compression;

namespace FileProcessingLibrary.Services
{
    public class ZipFileService : IZipFileService
    {
        public Dictionary<FileSource, ZipArchiveEntry> GetEntriesFromZipFile(ZipArchive archive)
        {
            //await using var ms = new MemoryStream();
            //await stream.CopyToAsync(ms);

            //using var archive = new ZipArchive(ms);
            var entries = archive.Entries;


            var masterFile = entries.FirstOrDefault(e => e.FullName.Contains("master", StringComparison.CurrentCultureIgnoreCase));
            var rbBacklogFile = entries.FirstOrDefault(e => e.FullName.Contains("rb", StringComparison.CurrentCultureIgnoreCase) && e.FullName.Contains("backlog", StringComparison.CurrentCultureIgnoreCase));
            var rbBillingFile = entries.FirstOrDefault(e => e.FullName.Contains("rb", StringComparison.CurrentCultureIgnoreCase) && e.FullName.Contains("billing", StringComparison.CurrentCultureIgnoreCase));
            var mfgFile = entries.FirstOrDefault(e => e.FullName.Contains("mfg", StringComparison.CurrentCultureIgnoreCase));

            

            if (masterFile is null)
            {
                return [];
            }

            var data = new Dictionary<FileSource, ZipArchiveEntry>
        {
            { FileSource.Master, masterFile },
            { FileSource.RbBacklog, rbBacklogFile },
            { FileSource.RbBilling, rbBillingFile },
            { FileSource.Mfg, mfgFile }
        };

            return data;
        }
    }
}

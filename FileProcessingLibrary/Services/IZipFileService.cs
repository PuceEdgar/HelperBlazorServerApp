using System.IO.Compression;

namespace FileProcessingLibrary.Services;

public interface IZipFileService
{
    Dictionary<FileSource, ZipArchiveEntry> GetEntriesFromZipFile(ZipArchive archive);
}

namespace FileProcessingLibrary;

public record PartExistsData
{
    public FileSource SourceData { get; set; }
    public string? PartNumber { get; set; }
}
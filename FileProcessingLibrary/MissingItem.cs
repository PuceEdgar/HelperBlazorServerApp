namespace FileProcessingLibrary;

public record MissingItem
{
    public string? PartNumber { get; set; }
    public string? CustomerNumber { get; set; }
    public int Qty { get; set; }
    public FileSource MissingIn { get; set; }
    public FileSource ExistsIn { get; set; }
}

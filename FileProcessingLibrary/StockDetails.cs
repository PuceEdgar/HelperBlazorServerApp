namespace FileProcessingLibrary;

public record StockDetails
{
    public string? CustomerNumber { get; set; }
    public string? PartNumber { get; set; }
    public int Qty { get; set; }
    public decimal UnitPrice { get; set; }
}

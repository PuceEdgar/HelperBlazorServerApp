namespace FileProcessingLibrary;

public record GroupedStockItem
{
    public FileSource StockSource { get; set; }
    public List<StockDetails> GroupedStockList { get; set; } = [];
}

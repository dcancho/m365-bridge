public class CellResult
{
	public CellAddress CellValue { get; set; }
	public string? Formula { get; set; }
	public string? Error { get; set; }
	public (CellAddress start, CellAddress end)? MergedRange { get; set; }
}
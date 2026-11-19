public class CellResult
{
	public CellAddress Address { get; set; }
	public string? Value { get; set; }
	public string? Formula { get; set; }
	public string? Error { get; set; }
	public (CellAddress start, CellAddress end)? MergedRange { get; set; }
}
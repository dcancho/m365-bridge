public class MatrixResult
{
	public CellAddress StartAddress { get; set; }
	public CellAddress EndAddress { get; set; }

	public List<List<CellResult>> Values { get; set; } = new List<List<CellResult>>();

	public CellResult? this[int row, int column]
	{
		get
		{
			if (row < 0 || row >= Values.Count)
				return null;

			var rowValues = Values[row];
			if (column < 0 || column >= rowValues.Count)
				return null;

			return rowValues[column];
		}
	}
}
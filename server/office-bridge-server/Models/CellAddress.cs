using System.Text.RegularExpressions;

public class CellAddress
{
	public int Row;
	public string Column;

	private CellAddress(int row, string column)
	{
		Row = row;
		Column = column;
	}

	public CellAddress? Set(int row, string column)
	{
		if (row <= 0 || row > 1_048_576 || GetColumnInt(column) == null )
		{
			return null;
		}

		return new CellAddress(row, column);
	}

	public static int? GetColumnInt(string column)
	{
		if (column.Length == 0) return null;
		if (column.Length > 3) return null;
		
		var pattern = new Regex(@"^[A-Z]{1,3}$");
		if (!pattern.IsMatch(column)) return null;
		
		int result = 0;
		for (int i = 0; i < column.Length; i++)
		{
			result *= 26;
			result += column[i] - 'A' + 1;
		}
		return result;
	}

	public static string? GetColumnString(int column)
	{
		if (column <= 0 || column > 16384) return null;

		string result = string.Empty;
		while (column > 0)
		{
			column--;
			result = (char)('A' + (column % 26)) + result;
			column /= 26;
		}
		return result;
	}
}
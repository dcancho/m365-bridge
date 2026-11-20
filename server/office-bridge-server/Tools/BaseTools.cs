using System.ComponentModel;
using System.Text.Json;
using Microsoft.AspNetCore.SignalR;
using ModelContextProtocol;
using ModelContextProtocol.Server;

internal class BaseTools(SessionRegistry registry, IHubContext<Relay> hub)
{
	private static readonly TimeSpan CommandTimeout = TimeSpan.FromSeconds(30);

	[McpServerTool(Name = "list_documents", ReadOnly = true)]
	[Description("Lists the Office documents that currently have the bridge add-in open. Call this first to get a document_id.")]
	public object ListDocuments()
	{
		return registry.GetSessions().Select(s => new
		{
			document_id = s.SessionId,
			host = s.Host,
			name = s.DocumentName,
			url = s.Url,
			opened_at = s.CreatedAt
		});
	}

	[McpServerTool(Name = "get_worksheets", ReadOnly = true)]
	[Description("Lists the worksheet names of an Excel workbook, in tab order.")]
	public Task<JsonElement?> GetWorksheets(
		[Description("Document id from list_documents.")] string document_id,
		CancellationToken cancellationToken = default)
	{
		return Execute(document_id, "get_worksheets", new { }, cancellationToken);
	}

	[McpServerTool(Name = "get_selection", ReadOnly = true)]
	[Description("Returns the sheet and address of the cells the user has selected in Excel. Use it when the user refers to \"this\", \"here\" or \"the selection\".")]
	public Task<JsonElement?> GetSelection(
		[Description("Document id from list_documents.")] string document_id,
		CancellationToken cancellationToken = default)
	{
		return Execute(document_id, "get_selection", new { }, cancellationToken);
	}

	[McpServerTool(Name = "read_range", ReadOnly = true)]
	[Description("Reads the values of a range in an Excel workbook. Returns a 2D array of values, one inner array per row, the formulas in the same shape when requested, and a read_token. write_range needs that token to change any cell inside this range.")]
	public Task<JsonElement?> ReadRange(
		[Description("Document id from list_documents.")] string document_id,
		[Description("Range address in A1 notation, such as \"A1\" or \"A1:C10\".")] string range,
		[Description("Worksheet name. Uses the active sheet when omitted.")] string? sheet = null,
		[Description("Also return the formula of each cell. Cells without a formula return their value.")] bool include_formulas = false,
		CancellationToken cancellationToken = default)
	{
		return Execute(document_id, "read_range", new { range, sheet, includeFormulas = include_formulas }, cancellationToken);
	}

	[McpServerTool(Name = "write_range", Destructive = true, Idempotent = true)]
	[Description("Writes values to a range in an Excel workbook, overwriting what is there. Requires the read_token from a read_range call whose range contains this one; it fails if those cells changed since the read, so read again and reconsider. The token stays valid after your own writes. The values array must have exactly the rows and columns of the range. Strings that start with \"=\" are written as formulas.")]
	public Task<JsonElement?> WriteRange(
		[Description("Document id from list_documents.")] string document_id,
		[Description("Range address in A1 notation, such as \"A1\" or \"A1:C10\".")] string range,
		[Description("Values to write, one inner array per row.")] JsonElement[][] values,
		[Description("read_token returned by read_range for a range that contains this one.")] string read_token,
		[Description("Worksheet name. Uses the active sheet when omitted.")] string? sheet = null,
		CancellationToken cancellationToken = default)
	{
		return Execute(document_id, "write_range", new { range, sheet, values, readToken = read_token }, cancellationToken);
	}

	private async Task<JsonElement?> Execute(string documentId, string command, object args, CancellationToken cancellationToken)
	{
		var session = registry.GetSession(documentId)
			?? throw new McpException($"Document '{documentId}' is not open. Call list_documents to see the open documents.");

		using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
		timeout.CancelAfter(CommandTimeout);

		ExcelResponse response;
		try
		{
			response = await hub.Clients.Client(session.ConnectionId)
				.InvokeAsync<ExcelResponse>("execute", command, args, timeout.Token);
		}
		catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
		{
			throw new McpException($"The add-in in '{session.DocumentName}' did not respond within {CommandTimeout.TotalSeconds} seconds. The command may still complete in Excel; read the range again before retrying a write.");
		}

		if (!response.Ok)
		{
			throw new McpException($"{response.Code}: {response.Message}");
		}

		return response.Data;
	}
}

public record ExcelResponse(bool Ok, JsonElement? Data, string? Code, string? Message);

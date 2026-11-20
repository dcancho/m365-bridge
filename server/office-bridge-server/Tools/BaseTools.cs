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

	[McpServerTool(Name = "read_range", ReadOnly = true)]
	[Description("Reads the values of a range in an Excel workbook. Returns a 2D array of values, one inner array per row.")]
	public Task<JsonElement?> ReadRange(
		[Description("Document id from list_documents.")] string document_id,
		[Description("Range address in A1 notation, such as \"A1\" or \"A1:C10\".")] string range,
		[Description("Worksheet name. Uses the active sheet when omitted.")] string? sheet = null,
		CancellationToken cancellationToken = default)
	{
		return Execute(document_id, "read_range", new { range, sheet }, cancellationToken);
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
			throw new McpException($"The add-in in '{session.DocumentName}' did not respond within {CommandTimeout.TotalSeconds} seconds.");
		}

		if (!response.Ok)
		{
			throw new McpException($"{response.Code}: {response.Message}");
		}

		return response.Data;
	}
}

public record ExcelResponse(bool Ok, JsonElement? Data, string? Code, string? Message);

using System.Collections.Concurrent;

public class SessionRegistry
{
	private readonly ConcurrentDictionary<string, Session> sessions = new();

	public void AddSession(RegisterRequest request, string connectionId)
	{
		sessions.AddOrUpdate(
			request.SessionId,
			_ => new Session
			{
				ConnectionId = connectionId,
				SessionId = request.SessionId,
				Host = request.Host,
				CreatedAt = DateTime.UtcNow,
				DocumentName = request.Name,
				Url = request.Url
			},
			(_, existing) => existing with
			{
				ConnectionId = connectionId,
				DocumentName = request.Name,
				Url = request.Url
			});
	}

	public void RemoveByConnection(string connectionId)
	{
		foreach (var session in sessions.Values.Where(s => s.ConnectionId == connectionId))
		{
			sessions.TryRemove(session.SessionId, out _);
		}
	}

	public Session? GetSession(string sessionId)
	{
		return sessions.TryGetValue(sessionId, out var session) ? session : null;
	}

	public List<Session> GetSessions()
	{
		return sessions.Values.ToList();
	}
}

public record Session
{
	public required string ConnectionId { get; init; }
	public required string SessionId { get; init; }
	public required string Host { get; init; }
	public required DateTime CreatedAt { get; init; }
	public required string DocumentName { get; init; }
	public required string Url { get; init; }
}

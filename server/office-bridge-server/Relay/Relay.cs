using Microsoft.AspNetCore.SignalR;

public class Relay(SessionRegistry registry) : Hub
{
	public async Task SendMessage(string method, object message)
	{
		await Clients.All.SendAsync(method, message);
	}

	public void Register(RegisterRequest request)
	{
		registry.AddSession(request, Context.ConnectionId);
	}

	public override Task OnDisconnectedAsync(Exception? exception)
	{
		registry.RemoveByConnection(Context.ConnectionId);
		return base.OnDisconnectedAsync(exception);
	}
}

public class RegisterRequest
{
	public required string SessionId { get; set; }
	public required string Host { get; set; }
	public string Name { get; set; } = "";
	public string Url { get; set; } = "";
}

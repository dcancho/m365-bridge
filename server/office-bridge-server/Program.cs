using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args = args,
    ContentRootPath = AppContext.BaseDirectory
});

builder.Services
    .AddMcpServer()
    .WithHttpTransport(options =>
    {
        // Stateless mode is recommended for servers that don't need
        // server-to-client requests like sampling or elicitation.
        // See https://csharp.sdk.modelcontextprotocol.io/concepts/transports/transports.html for details.
        options.Stateless = true;
    })
    .WithTools<BaseTools>();

builder.Services.AddSignalR();

builder.Services.AddSingleton<SessionRegistry>();

var addinOrigins = builder.Configuration.GetSection("AddinOrigins").Get<string[]>() ?? [];
builder.Services.AddCors(options =>
{
    options.AddPolicy("Addin", policy => policy
        .WithOrigins(addinOrigins)
        .AllowAnyHeader()
        .AllowAnyMethod()
        .AllowCredentials());
});

var app = builder.Build();

var addinPath = Path.Combine(AppContext.BaseDirectory, app.Configuration["AddinPath"] ?? "addin");
if (Directory.Exists(addinPath))
{
    app.UseStaticFiles(new StaticFileOptions { FileProvider = new PhysicalFileProvider(addinPath) });
}

app.UseCors();
app.MapMcp();
app.MapHub<Relay>("/relay").RequireCors("Addin");

app.Run();

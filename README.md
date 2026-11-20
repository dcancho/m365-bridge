# m365-bridge

A relay that lets an AI agent read and write a live Excel workbook. An Office add-in opens an outbound connection to the server, and the server exposes the workbook to agents as MCP tools.

## Layout

- `add-in/`: Excel task pane add-in (TypeScript, Office.js)
- `server/office-bridge-server/`: relay and MCP server (ASP.NET Core)
- `scripts/`: publish, install and uninstall on Windows

## Installing on Windows

Requires the .NET SDK, Node.js and a trusted HTTPS development certificate (`dotnet dev-certs https --trust`).

```powershell
.\scripts\publish.ps1
.\scripts\install.ps1
```

Then, once, from an administrator PowerShell:

```powershell
.\scripts\share-catalog.ps1
```

`publish.ps1` builds the add-in and publishes the server, with the add-in inside it, to `publish/`. `install.ps1` copies the manifest to `publish/catalog`, registers that folder as a trusted add-in catalog in Office, starts the server at sign-in, and starts it now. `share-catalog.ps1` shares the catalog folder, since Office only reads catalogs from network paths. The server runs without a window and listens on:

- `https://localhost:5299`: add-in files and the relay hub
- `http://localhost:6143`: MCP endpoint

Restart Excel and add the add-in once from **Add-ins > My Add-ins > Shared Folder > M365 Bridge**. After that, open the pane from **Home > M365 Bridge > Bridge** in any workbook. To remove everything, run `.\scripts\uninstall.ps1`.

## Connecting Claude Desktop

Claude Desktop starts MCP servers over stdio, so it reaches the HTTP endpoint through [`mcp-remote`](https://www.npmjs.com/package/mcp-remote). Add this to `claude_desktop_config.json` (Settings > Developer > Edit Config) and restart Claude Desktop:

```json
{
  "mcpServers": {
    "m365-bridge": {
      "command": "npx",
      "args": ["-y", "mcp-remote", "http://localhost:6143"]
    }
  }
}
```

## Developing

Add-in, served on `https://localhost:3100` and connected to the server on `https://localhost:5299`:

```bash
cd add-in
npm install
npm start
```

`npm start` sideloads the add-in into Excel desktop as M365 Bridge Dev, next to the installed one. `npm run start:debug` does the same with the WebView debugger attached.

Server:

```bash
dotnet run --project server/office-bridge-server --launch-profile https
```

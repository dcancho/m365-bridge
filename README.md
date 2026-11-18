# m365-bridge

A relay that lets an AI agent read and write a live Excel workbook. An Office add-in opens an outbound connection to the server, and the server exposes the workbook to agents as MCP tools.

## Layout

- `add-in/`: Excel task pane add-in (TypeScript, Office.js)
- `server/office-bridge-server/`: relay and MCP server (ASP.NET Core)

## Running locally

Add-in, served on `https://localhost:3100`:

```bash
cd add-in
npm install
npm start
```

`npm start` sideloads the add-in into Excel desktop. `npm run start:debug` does the same with the WebView debugger attached.

Server, MCP endpoint on `http://localhost:6143`:

```bash
dotnet run --project server/office-bridge-server
```

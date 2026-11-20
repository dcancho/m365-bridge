import { runExcel, getDocumentName } from "./excel";
import * as signalR from "@microsoft/signalr";

declare const RELAY_URL: string;

export type RegisterRequest = {
	sessionId: string;
	host: string;
	name: string;
	url: string;
}

type Status = "connecting" | "connected" | "disconnected";

const sessionId = crypto.randomUUID();
const retryDelayMs = 5000;

const connection = new signalR.HubConnectionBuilder()
	.withUrl(RELAY_URL)
	.withAutomaticReconnect()
	.configureLogging(signalR.LogLevel.Information)
	.build();

connection.on("execute", (commandName: string, args: unknown) => runExcel(commandName, args));

function setStatus(status: Status, text: string) {
	document.getElementById("status")!.dataset.state = status;
	document.getElementById("status-text")!.textContent = text;
}

async function register() {
	if (connection.state !== signalR.HubConnectionState.Connected) {
		return;
	}

	const request: RegisterRequest = {
		sessionId,
		host: String(Office.context.host),
		name: await getDocumentName().catch(() => ""),
		url: Office.context.document.url ?? "",
	};

	try {
		await connection.invoke("register", request);
		setStatus("connected", "Connected");
	} catch (err) {
		console.error(err);
		setStatus("disconnected", "Registration failed");
		setTimeout(register, retryDelayMs);
	}
}

async function start() {
	setStatus("connecting", "Connecting");
	try {
		await connection.start();
	} catch (err) {
		console.error(err);
		setStatus("disconnected", "Bridge server unreachable");
		setTimeout(start, retryDelayMs);
		return;
	}
	await register();
}

connection.onreconnecting(() => setStatus("connecting", "Reconnecting"));
connection.onreconnected(() => register());
connection.onclose(() => start());

Office.onReady(() => start());

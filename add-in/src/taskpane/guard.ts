// Read tokens: a write is only accepted for cells the agent has read and that
// nobody has changed since. The panel holds the state because it is the one
// piece that lives exactly as long as the document is open.

type ReadEntry = {
	worksheetId: string;
	address: string;
	hash: string;
};

const MAX_TOKENS = 200;
const tokens = new Map<string, ReadEntry>();

export class BridgeError extends Error {
	constructor(public code: string, message: string) {
		super(message);
	}
}

// FNV-1a, 32 bits. It only has to notice changes, not resist an attacker:
// the agent never sees the hash, only the random token that points to it.
export function hashCells(formulas: unknown[][]): string {
	const text = JSON.stringify(formulas);
	let h = 0x811c9dc5;
	for (let i = 0; i < text.length; i++) {
		h ^= text.charCodeAt(i);
		h = Math.imul(h, 0x01000193);
	}
	return (h >>> 0).toString(16).padStart(8, "0");
}

function newToken(): string {
	const alphabet = "abcdefghjkmnpqrstuvwxyz23456789";
	const bytes = crypto.getRandomValues(new Uint8Array(7));
	return Array.from(bytes, (b) => alphabet[b % alphabet.length]).join("");
}

export function issueToken(entry: ReadEntry): string {
	const token = newToken();
	tokens.set(token, entry);
	if (tokens.size > MAX_TOKENS) {
		tokens.delete(tokens.keys().next().value!);
	}
	return token;
}

export function getToken(token: string): ReadEntry {
	const entry = tokens.get(token);
	if (!entry) {
		throw new BridgeError(
			"ReadRequired",
			`Read token '${token}' is unknown or expired. Call read_range on the cells you want to change and use the token it returns.`
		);
	}
	return entry;
}

export function refreshToken(token: string, hash: string) {
	const entry = tokens.get(token);
	if (entry) entry.hash = hash;
}

import { BridgeError, getToken, hashCells, issueToken, refreshToken } from "./guard";

type Handler = (ctx: Excel.RequestContext, args: any) => Promise<unknown>;

const handlers: Record<string, Handler> = {
	"read_range": readRange,
	"get_worksheets": getWorksheets,
	"get_selection": getSelection,
	"write_range": writeRange,
};

// A full column (A:A) is a million cells. Refuse before loading it.
const MAX_CELLS = 20000;

export interface Response {
	ok: boolean;
	data?: any;
	code?: string;
	message?: string;
}

export async function getDocumentName(): Promise<string> {
	return Excel.run(async (ctx) => {
		ctx.workbook.load("name");
		await ctx.sync();
		return ctx.workbook.name;
	});
}

export async function runExcel(commandName: string, args: any): Promise<Response> {
	try {
		if (!Object.hasOwn(handlers, commandName)) {
			return {
				ok: false,
				code: "UnknownCommand",
				message: `Command ${commandName} not found`,
			};
		}

		return {
			ok: true,
			data: await Excel.run((ctx: Excel.RequestContext) => handlers[commandName](ctx, args)),
		};
	} catch (error) {
		if (error instanceof BridgeError || error instanceof OfficeExtension.Error) {
			return { ok: false, code: error.code, message: error.message };
		}
		return { ok: false, code: "UnknownError", message: String(error) };
	}
}

function getWorksheet(ctx: Excel.RequestContext, sheet?: string): Excel.Worksheet {
	return sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();
}

async function checkSize(ctx: Excel.RequestContext, range: Excel.Range) {
	range.load("cellCount");
	await ctx.sync();
	if (range.cellCount > MAX_CELLS) {
		throw new BridgeError(
			"RangeTooLarge",
			`The range has ${range.cellCount} cells; the limit is ${MAX_CELLS}. Read it in smaller blocks.`
		);
	}
}

export type ReadRangeArgs = {
	range: string;
	sheet?: string;
	includeFormulas?: boolean;
};

export async function readRange(
	ctx: Excel.RequestContext,
	{ range, sheet, includeFormulas = false }: ReadRangeArgs
) {
	const ws = getWorksheet(ctx, sheet);
	const rangeObject = ws.getRange(range);
	await checkSize(ctx, rangeObject);

	ws.load("id, name");
	rangeObject.load("address, values, formulas");
	await ctx.sync();

	// Hash formulas, not values: a volatile formula like NOW() recalculates
	// without anyone touching the sheet and would invalidate every token.
	const token = issueToken({
		worksheetId: ws.id,
		address: rangeObject.address,
		hash: hashCells(rangeObject.formulas),
	});

	return {
		sheet: ws.name,
		address: rangeObject.address,
		values: rangeObject.values,
		...(includeFormulas ? { formulas: rangeObject.formulas } : {}),
		read_token: token,
	};
}

export async function getWorksheets(ctx: Excel.RequestContext): Promise<string[]> {
	ctx.workbook.worksheets.load("items/name");
	await ctx.sync();
	return ctx.workbook.worksheets.items.map((ws) => ws.name);
}

export async function getSelection(ctx: Excel.RequestContext) {
	const selection = ctx.workbook.getSelectedRange();
	const ws = selection.worksheet;
	selection.load("address, rowCount, columnCount");
	ws.load("name");
	await ctx.sync();

	return {
		sheet: ws.name,
		address: selection.address,
		rows: selection.rowCount,
		columns: selection.columnCount,
	};
}

export type WriteRangeArgs = {
	values: any[][];
	range: string;
	sheet?: string;
	readToken: string;
};

export async function writeRange(
	ctx: Excel.RequestContext,
	{ values, range, sheet, readToken }: WriteRangeArgs
) {
	const entry = getToken(readToken);

	const ws = getWorksheet(ctx, sheet);
	const target = ws.getRange(range);
	const readArea = ctx.workbook.worksheets.getItem(entry.worksheetId).getRange(entry.address);
	// The bounding box of read ∪ target equals the read area only when the
	// target lies inside it.
	const bounds = readArea.getBoundingRect(target);

	ws.load("id");
	target.load("address");
	bounds.load("address");
	readArea.load("formulas");
	await ctx.sync();

	if (ws.id !== entry.worksheetId || bounds.address !== entry.address) {
		throw new BridgeError(
			"OutsideRead",
			`Token '${readToken}' covers ${entry.address}, which does not contain ${target.address}. Read the cells you want to write first.`
		);
	}

	if (hashCells(readArea.formulas) !== entry.hash) {
		throw new BridgeError(
			"StaleRead",
			`${entry.address} changed since it was read. Read it again before writing.`
		);
	}

	target.values = values;
	readArea.load("formulas");
	await ctx.sync();

	// The change is the agent's own, so the token stays valid for the next write.
	refreshToken(readToken, hashCells(readArea.formulas));

	return { address: target.address, read_token: readToken };
}

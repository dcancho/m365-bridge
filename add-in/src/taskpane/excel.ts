type Handler = (ctx: Excel.RequestContext, args: any) => Promise<unknown>

const handlers: Record<string, Handler> = {
	"read_range": readRange,
	"get_worksheets": getWorksheets,
	"write_range": writeRange
}

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
				message: `Command ${commandName} not found`
			}
		}

		return {
			ok: true,
			data: await Excel.run((ctx: Excel.RequestContext) => handlers[commandName](ctx, args))
		};
	} catch (error) {
		if (error instanceof OfficeExtension.Error) {
			return { ok: false, code: error.code, message: error.message };
		}
		return { ok: false, code: "UnknownError", message: String(error) };
	}
}

export type ReadRangeArgs = {
	range: string;
	sheet?: string;
	includeFormulas?: boolean;
}

export async function readRange(
	ctx: Excel.RequestContext,
	{ range, sheet, includeFormulas = false }: ReadRangeArgs
): Promise<{
	values: any[][];
	formulas?: any[][];
}> {
	const ws = sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();

	const rangeObject = ws.getRange(range);

	rangeObject.load("values");
	if (includeFormulas) rangeObject.load("formulas");

	await ctx.sync();

	return includeFormulas
		? { values: rangeObject.values, formulas: rangeObject.formulas }
		: { values: rangeObject.values };
}

export async function getWorksheets(
	ctx: Excel.RequestContext
): Promise<string[]> {
	ctx.workbook.worksheets.load("items");
	await ctx.sync();
	return ctx.workbook.worksheets.items.map((ws) => ws.name);
}

export type WriteRangeArgs = {
	values: any[][];
	range: string;
	sheet?: string;
}

export async function writeRange(
	ctx: Excel.RequestContext,
	{ values, range, sheet }: WriteRangeArgs
): Promise<void> {
	const ws = sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();

	const rangeObject = ws.getRange(range);
	rangeObject.values = values;
	await ctx.sync();
}


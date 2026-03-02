import * as extensionConfig from '../extension.json';
import * as ClipperLib from 'clipper-lib';

const CHANNEL = 'KLE_IMPORTER_PRO';
let messageListenerInstalled = false;

function showInfo(message: unknown, title = 'KLE Importer Pro') {
	const msg = typeof message === 'string' ? message : String(message);
	if (eda?.sys_Dialog?.showInformationMessage) {
		eda.sys_Dialog.showInformationMessage(msg, title);
	}
	else if (eda?.sys_MessageBox?.showInformationMessage) {
		eda.sys_MessageBox.showInformationMessage(msg, title);
	}
}

function showToast(message: unknown, seconds = 2) {
	const msg = typeof message === 'string' ? message : String(message);
	if (eda?.sys_Message?.showToastMessage) {
		eda.sys_Message.showToastMessage(msg, 'info' as any, seconds);
		return;
	}
	if (eda?.sys_ToastMessage?.showMessage) {
		eda.sys_ToastMessage.showMessage(msg, seconds);
	}
}

function getSysIFrame() {
	return (eda as any)?.sys_IFrame ?? (eda as any)?.sys_iframe ?? null;
}

function hideUi() {
	try {
		const sysIFrame = getSysIFrame();
		if (sysIFrame?.hideIFrame) sysIFrame.hideIFrame('kle-importer-pro');
	}
	catch {}
}

function ensureMessageListener() {
	if (messageListenerInstalled) return;
	if (typeof window === 'undefined' || !window.addEventListener) return;

	const handler = (event: MessageEvent) => {
		const data: any = (event as any)?.data;
		if (!data || data.channel !== CHANNEL) return;

		const reply = (payload: any) => {
			try {
				const src = (event as any)?.source;
				if (src?.postMessage) src.postMessage(payload, '*');
			}
			catch {}
		};

		if (data.type === 'HELLO') {
			reply({ channel: CHANNEL, type: 'HELLO_ACK', token: data.token, ok: true });
			return;
		}

		if (data.type === 'PICK_FILE') {
			void (async () => {
				try {
					showToast('KLE Importer Pro: opening file dialog...', 2);
					const token = data.token;
					const fsApi: any =
						(eda as any)?.sys_FileSystem
						?? (eda as any)?.sys_fileSystem
						?? (eda as any)?.sys_filesystem
						?? (eda as any)?.sys_FileSystem;

					if (!fsApi?.openReadFileDialog) {
						reply({
							channel: CHANNEL,
							type: 'FILE_PICKED',
							token,
							ok: false,
							error: 'sys_FileSystem.openReadFileDialog not available',
						}, '*');
						showInfo('sys_FileSystem.openReadFileDialog が利用できません。', 'Error');
						return;
					}

					const res = await fsApi.openReadFileDialog();
					if (!res) {
						reply({ channel: CHANNEL, type: 'FILE_PICKED', token, ok: false, error: 'cancelled' });
						return;
					}
					const file = Array.isArray(res) ? res[0] : res;
					const text = await file.text();
					const name = file.name || '';
					reply({ channel: CHANNEL, type: 'FILE_PICKED', token, ok: true, payload: { name, text } });
				}
				catch (e: any) {
					reply({ channel: CHANNEL, type: 'FILE_PICKED', token: data.token, ok: false, error: e?.message ?? String(e) });
				}
			})();
		}

		if (data.type === 'RUN') {
			const p = data.payload || {};
			showToast('KLE Importer Pro: RUN received', 2);
			reply({ channel: CHANNEL, type: 'RUN_ACK', token: data.token, ok: true });
			hideUi();
			void applyPlacement(p as PlacementPayload);
		}
		if (data.type === 'CLOSE') {
			hideUi();
		}
	};

	// Register on multiple window references (some EasyEDA frames may route messages differently).
	window.addEventListener('message', handler);
	try {
		if (window.top && window.top !== window) window.top.addEventListener('message', handler);
	}
	catch {}
	try {
		if (window.parent && window.parent !== window) window.parent.addEventListener('message', handler);
	}
	catch {}
	messageListenerInstalled = true;
}

async function openDialog() {
	const sysIFrame = getSysIFrame();
	if (!sysIFrame?.openIFrame) {
		showInfo('eda.sys_IFrame.openIFrame が利用できません。', 'Error');
		return;
	}

	hideUi();
	try {
		await sysIFrame.openIFrame('/iframe/index.html', 380, 610, 'kle-importer-pro', {
			grayscaleMask: false,
			maximizeButton: false,
			minimizeButton: false,
		});
	}
	catch (e: any) {
		showInfo(`openIFrame エラー: ${e?.message ?? e}`, 'Error');
	}
}

export function activate(_status?: 'onStartupFinished', _arg?: string): void {
	// Keep startup quiet (no dialogs/toasts). If needed, use the menu "Test ping".
	ensureMessageListener();

	// Expose a direct-call bridge for iframe pages (more reliable than postMessage in some EasyEDA builds).
	try {
		(eda as any).jlc_eda_kle_importer_pickFile = async () => {
			const fsApi: any =
				(eda as any)?.sys_FileSystem
				?? (eda as any)?.sys_fileSystem
				?? (eda as any)?.sys_filesystem;
			if (!fsApi?.openReadFileDialog) {
				throw new Error('sys_FileSystem.openReadFileDialog not available');
			}
			const res = await fsApi.openReadFileDialog();
			if (!res) return null;
			const file = Array.isArray(res) ? res[0] : res;
			return { name: file?.name || '', text: await file.text() };
		};
		(eda as any).jlc_eda_kle_importer_runPlacement = async (payload: any) => {
			await applyPlacement(payload || {});
			return true;
		};
		(eda as any).jlc_eda_kle_importer_exportPlate = async (payload: any) => {
			return await exportSwitchPlate(payload || {});
		};
		(eda as any).jlc_eda_kle_importer_exportCase = async (payload: any) => {
			return await exportCase(payload || {});
		};
		(eda as any).jlc_eda_kle_importer_closeUi = () => closeUi();
	}
	catch {}
}

export function ping(): void {
	showInfo('KLE Importer Pro: ping OK');
}

export function closeUi(): void {
	hideUi();
	showToast('KLE Importer Pro: UI closed', 2);
}

export function run(): void {
	ensureMessageListener();
	showToast('KLE Importer Pro: opening...', 2);
	void openDialog();
}

type OrderMode = 'legend' | 'kle' | 'y_x' | 'x_y' | 'custom';

type PlacementPayload = {
	json: string;
	pitch: number;
	dx: number;
	dy: number;
	doDiode: boolean;
	// Backward/forward compatibility: ignore extra fields from older UI versions.
	[key: string]: unknown;
};

type PlateExportPayload = {
	json: string;
	pitch: number; // mm per U
	format: 'svg' | 'dxf';
	marginMm: number;
	cutoutSizeMm: number;
	applyKeyRotation: boolean;
	outlineMode: 'hull' | 'rectUnion';
	rectUnionMarginMm: number; // used when outlineMode=rectUnion
	[key: string]: unknown;
};

type CaseExportPayload = {
	json: string;
	pitch: number; // mm per U
	format: 'svg' | 'dxf';
	boardClearanceMm: number; // offset from PCB outline (e.g. 0.2)
	wallThicknessMm: number; // additional offset for case outer (e.g. 2.0)
	includeKeyHoles: boolean;
	keyHoleSizeMm: number;
	applyKeyRotation: boolean;
	outlineSource: 'board' | 'selected' | 'plateHull' | 'rectUnion'; // default 'board'
	// When outlineSource is 'plateHull' or 'rectUnion', build outline from SW footprint centers.
	outlineRectSizeMm?: number; // default pitch
	rectUnionMarginMm?: number; // default 2.0 (for rectUnion/plateHull)
	[key: string]: unknown;
};

function stripBomAndTrim(text: string): string {
	// JSON.parse/new Function in some environments choke on a leading BOM.
	return text.replace(/^\uFEFF/, '').trim();
}

function previewText(text: string, maxLen = 120): string {
	const t = text.length > maxLen ? text.slice(0, maxLen) + '…' : text;
	// Make control chars visible-ish
	return t.replace(/[\u0000-\u001F\u007F]/g, (c) => {
		const code = c.charCodeAt(0).toString(16).padStart(2, '0');
		return `\\x${code}`;
	});
}

function parseKleText(raw: unknown): any {
	if (typeof raw !== 'string') {
		throw new Error(`KLE JSON is not a string (type=${typeof raw})`);
	}
	const original = raw;
	let text = stripBomAndTrim(original);

	const tryJson = (s: string) => JSON.parse(s);
	const tryJs = (s: string) => new Function(`return (${s})`)();

	let jsonErr: any = null;
	try {
		return tryJson(text);
	}
	catch (e: any) {
		jsonErr = e;
	}

	// Some KLE exports are wrapped or have extra prefix/suffix; try extracting the array.
	const firstArr = text.indexOf('[');
	const lastArr = text.lastIndexOf(']');
	if (firstArr !== -1 && lastArr !== -1 && lastArr > firstArr) {
		const sliced = stripBomAndTrim(text.slice(firstArr, lastArr + 1));
		try {
			return tryJson(sliced);
		}
		catch {}
	}

	let jsErr: any = null;
	try {
		return tryJs(text);
	}
	catch (e: any) {
		jsErr = e;
	}

	const parts: string[] = [];
	if (jsonErr?.message) parts.push(`JSON.parse: ${jsonErr.message}`);
	if (jsErr?.message) parts.push(`JS eval: ${jsErr.message}`);
	parts.push(`preview: ${previewText(text)}`);
	parts.push(`length: ${text.length}`);
	throw new Error(parts.join('\n'));
}

function parseSwHint(label: unknown): number | null {
	if (typeof label !== 'string') return null;
	const m = /\bSW\s*([0-9]+)\b/i.exec(label);
	if (!m) return null;
	const n = parseInt(m[1], 10);
	return Number.isFinite(n) ? n : null;
}

function assignByLegend(layout: any[]): Array<{ key: any; n: number }> {
	const out: Array<{ key: any; n: number }> = [];
	for (const k of layout) {
		const hint = k?.swHint;
		if (typeof hint === 'number' && Number.isFinite(hint)) {
			out.push({ key: k, n: hint });
		}
	}
	return out;
}

async function applyPlacement(payload: PlacementPayload) {
	try {
		const rawJson = payload?.json;
		if (!rawJson) return;
		if (!eda?.pcb_PrimitiveComponent?.getAll || !eda?.pcb_PrimitiveComponent?.modify) {
			showInfo('PCB API が利用できません。PCBエディタを開いてから実行してください。', 'Error');
			return;
		}

		const pitch = Number(payload.pitch);
		const dxMm = Number(payload.dx);
		const dyMm = Number(payload.dy);
		const doDiode = Boolean(payload.doDiode);

		// Place strictly by legend "SWxx" mapping:
		// - No sequential assignment
		// - No sorting modes
		// - Keys without SWxx are ignored
		const orderMode: OrderMode = 'legend';

		const kle = parseKleText(rawJson);

		const pitchNum = Number(pitch);
		if (!Number.isFinite(pitchNum) || pitchNum <= 0) throw new Error('Invalid pitch');

		const milPerU = (pitchNum / 25.4) * 1000;
		const duX = (Number(dxMm) || 0) / pitchNum;
		const duY = (Number(dyMm) || 0) / pitchNum;

		showToast('配置を適用中です...', 2);

		const layout = parseKLE(kle, duX, duY);
		const allCompsRaw = await eda.pcb_PrimitiveComponent.getAll();
		const pcbComponents = Array.isArray(allCompsRaw) ? allCompsRaw : Object.values(allCompsRaw || {});

		let swCount = 0;
		let dCount = 0;

		if (orderMode === 'legend') {
			const assigned = assignByLegend(layout);
			if (assigned.length === 0) {
				showInfo('レジェンドに SWxx が見つかりませんでした。\n例: "...\\nSW4" のように SW番号を入れてください。', 'Error');
				return;
			}

			const seen = new Map<number, number>();
			for (const a of assigned) {
				seen.set(a.n, (seen.get(a.n) || 0) + 1);
			}
			const dups = [...seen.entries()].filter(([, c]) => c > 1).map(([n]) => n);
			if (dups.length) {
				showInfo(`レジェンドに重複したSW番号があります: ${dups.slice(0, 20).join(', ')}`, 'Warning');
			}

			for (const { key: k, n } of assigned) {
				const swDesignator = `SW${n}`;

				const sw = pcbComponents.find((c: any) => c?.designator === swDesignator);
				if (sw?.primitiveId) {
					await eda.pcb_PrimitiveComponent.modify(sw.primitiveId, {
						x: k.swX * milPerU,
						y: -k.swY * milPerU,
						rotation: -k.rot,
					});
					swCount++;
				}

				if (doDiode) {
					const dDesignator = `D${n}`;
					const d = pcbComponents.find((c: any) => c?.designator === dDesignator);
					if (d?.primitiveId) {
						await eda.pcb_PrimitiveComponent.modify(d.primitiveId, {
							x: k.dX * milPerU,
							y: -k.dY * milPerU,
							rotation: -k.rot,
						});
						dCount++;
					}
				}
			}
		}

		showInfo(`配置完了！\nSW: ${swCount}個\nDiode: ${dCount}個\n(ピッチ: ${pitchNum}mm)`);
	}
	catch (err: any) {
		showInfo(`エラー: ${err?.message ?? err}`, 'Error');
	}
}

function parseKLE(json: any[], duX: number, duY: number) {
	const result: Array<any> = [];
	let rx = 0;
	let ry = 0;
	let x = 0;
	let y = 0;
	let r = 0;

	json.forEach((row) => {
		if (!Array.isArray(row)) return;
		x = rx;
		let w = 1;
		let h = 1;

		row.forEach((item) => {
			if (item && typeof item === 'object' && !Array.isArray(item)) {
				if (item.r !== undefined) r = item.r;
				if (item.rx !== undefined) { rx = item.rx; x = rx; y = ry; }
				if (item.ry !== undefined) { ry = item.ry; y = ry; }
				if (item.w !== undefined) w = item.w;
				if (item.h !== undefined) h = item.h;
				if (item.x !== undefined) x += item.x;
				if (item.y !== undefined) y += item.y;
			}
			else if (typeof item === 'string') {
				const label = item;
				const cx = x + w / 2;
				const cy = y + h / 2;
				const rad = (r * Math.PI) / 180;

				const rotate = (px: number, py: number) => {
					const dx = px - rx;
					const dy = py - ry;
					return {
						x: rx + (dx * Math.cos(rad) - dy * Math.sin(rad)),
						y: ry + (dx * Math.sin(rad) + dy * Math.cos(rad)),
					};
				};

				const swPos = rotate(cx, cy);
				const dPos = rotate(cx + duX, cy + duY);
				result.push({
					label,
					swHint: parseSwHint(label),
					swX: swPos.x,
					swY: swPos.y,
					dX: dPos.x,
					dY: dPos.y,
					rot: r,
				});
				x += w;
				w = 1;
				h = 1;
			}
		});

		y += 1;
	});

	return result;
}

function mmToMil(mm: number): number {
	return (mm / 25.4) * 1000;
}

function asArray<T>(maybeArrayOrObj: any): T[] {
	if (!maybeArrayOrObj) return [];
	if (Array.isArray(maybeArrayOrObj)) return maybeArrayOrObj as T[];
	if (typeof maybeArrayOrObj === 'object') return Object.values(maybeArrayOrObj) as T[];
	return [];
}

function getPrimitiveIdFromUnknown(res: any): string | null {
	if (!res) return null;
	if (typeof res === 'string') return res;
	if (res.id) return String(res.id);
	if (res.primitiveId) return String(res.primitiveId);
	if (typeof res.getState_PrimitiveId === 'function') return String(res.getState_PrimitiveId());
	return null;
}

type Pt = { x: number; y: number };

function normalizePoints(pts: Pt[], intMil = true): Pt[] {
	// Remove NaNs, round (optional), and drop consecutive duplicates.
	const out: Pt[] = [];
	for (const p of pts || []) {
		let x = Number(p?.x);
		let y = Number(p?.y);
		if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
		if (intMil) {
			x = Math.round(x);
			y = Math.round(y);
		}
		const prev = out[out.length - 1];
		if (prev && prev.x === x && prev.y === y) continue;
		out.push({ x, y });
	}
	// Drop closing duplicate point if present.
	if (out.length >= 2) {
		const a = out[0];
		const b = out[out.length - 1];
		if (a.x === b.x && a.y === b.y) out.pop();
	}
	return out;
}

async function createPolygonRobust(mathApi: any, ptsIn: Pt[]): Promise<any | null> {
	if (!mathApi?.createPolygon) return null;
	const pts = normalizePoints(ptsIn, true);
	if (pts.length < 3) return null;

	const x0 = pts[0].x;
	const y0 = pts[0].y;

	const sources: any[] = [];

	// Variant A: old style "[x0,y0,'L', x1,y1, x2,y2, ...]"
	{
		const s: any[] = [x0, y0, 'L'];
		for (let i = 1; i < pts.length; i++) s.push(pts[i].x, pts[i].y);
		// close by repeating start
		s.push(x0, y0);
		sources.push(s);
	}

	// Variant B: explicit commands "[x0,y0,'L', x1,y1,'L',..., 'Z']"
	{
		const s: any[] = [x0, y0];
		for (let i = 1; i < pts.length; i++) s.push('L', pts[i].x, pts[i].y);
		s.push('Z');
		sources.push(s);
	}

	// Variant C: explicit commands with close by point, no Z
	{
		const s: any[] = [x0, y0];
		for (let i = 1; i < pts.length; i++) s.push('L', pts[i].x, pts[i].y);
		s.push('L', x0, y0);
		sources.push(s);
	}

	// Variant D: string path (some builds accept SVG-like commands)
	{
		const d = ['M', x0, y0, ...pts.slice(1).flatMap((p) => ['L', p.x, p.y]), 'Z'].join(' ');
		sources.push(d);
	}

	for (const src of sources) {
		try {
			const poly = await mathApi.createPolygon(src);
			if (poly) return poly;
		}
		catch {}
	}

	return null;
}

function getLayerId(p: any): number | null {
	let v =
		(typeof p?.getState_LayerId === 'function' ? p.getState_LayerId() : undefined)
		?? (typeof p?.getState_Layer === 'function' ? p.getState_Layer() : undefined)
		?? p?.layerId
		?? p?.layer
		?? p?.l
		?? p?.layer_id
		?? p?.layerID;
	const n = Number(v);
	return Number.isFinite(n) ? n : null;
}

function parsePointsFromString(s: string): Pt[] {
	// Be permissive: extract all numbers and treat them as x1,y1,x2,y2...
	// Works for formats like: "1,2 3,4" or "1 2,3 4" etc.
	const nums = (s.match(/-?\d+(?:\.\d+)?/g) || []).map((n) => Number(n)).filter((n) => Number.isFinite(n));
	const pts: Pt[] = [];
	for (let i = 0; i + 1 < nums.length; i += 2) {
		pts.push({ x: nums[i], y: nums[i + 1] });
	}
	return pts;
}

function extractPolylinePoints(poly: any): Pt[] | null {
	try {
		if (typeof poly?.getState_Polygon === 'function') {
			const polygon = poly.getState_Polygon();
			const arr = polygon?.polygon;
			if (Array.isArray(arr)) {
				const nums = arr.map((n: any) => Number(n)).filter((n: number) => Number.isFinite(n));
				const pts: Pt[] = [];
				for (let i = 0; i + 1 < nums.length; i += 2) {
					pts.push({ x: nums[i], y: nums[i + 1] });
				}
				if (pts.length >= 3) return pts;
			}
		}

		const ptsAny =
			poly?.points
			?? poly?.pointArr
			?? poly?.pointArray
			?? poly?.path
			?? poly?.data;

		if (typeof ptsAny === 'string') {
			const pts = parsePointsFromString(ptsAny);
			return pts.length >= 3 ? pts : null;
		}

		if (Array.isArray(ptsAny)) {
			const pts = ptsAny.map((p: any) => {
				const x = Number(p?.x ?? p?.X ?? p?.[0]);
				const y = Number(p?.y ?? p?.Y ?? p?.[1]);
				return { x, y };
			}).filter((p: Pt) => Number.isFinite(p.x) && Number.isFinite(p.y));
			return pts.length >= 3 ? pts : null;
		}

		// Some primitives expose a method returning points.
		if (typeof poly?.getState_Points === 'function') {
			const ret = poly.getState_Points();
			if (typeof ret === 'string') return extractPolylinePoints({ points: ret });
			if (Array.isArray(ret)) return extractPolylinePoints({ points: ret });
		}
	}
	catch {}
	return null;
}

function polygonArea(pts: Pt[]): number {
	let a = 0;
	for (let i = 0; i < pts.length; i++) {
		const p = pts[i];
		const q = pts[(i + 1) % pts.length];
		a += p.x * q.y - q.x * p.y;
	}
	return a / 2;
}

function bboxArea(pts: Pt[]): number {
	let minX = Infinity;
	let minY = Infinity;
	let maxX = -Infinity;
	let maxY = -Infinity;
	for (const p of pts) {
		minX = Math.min(minX, p.x);
		minY = Math.min(minY, p.y);
		maxX = Math.max(maxX, p.x);
		maxY = Math.max(maxY, p.y);
	}
	if (!Number.isFinite(minX)) return 0;
	return Math.max(0, maxX - minX) * Math.max(0, maxY - minY);
}

function pointInPolygon(pt: Pt, poly: Pt[]): boolean {
	// Ray casting. Works for simple polygons.
	let inside = false;
	for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
		const xi = poly[i].x, yi = poly[i].y;
		const xj = poly[j].x, yj = poly[j].y;
		const intersect = ((yi > pt.y) !== (yj > pt.y)) && (pt.x < ((xj - xi) * (pt.y - yi)) / (yj - yi + 0.0) + xi);
		if (intersect) inside = !inside;
	}
	return inside;
}

function extractComponentPoint(comp: any): Pt | null {
	const x = Number(
		(typeof comp?.getState_X === 'function' ? comp.getState_X() : undefined)
		?? (typeof comp?.getState_CenterX === 'function' ? comp.getState_CenterX() : undefined)
		?? comp?.x ?? comp?.X ?? comp?.centerX ?? comp?.cx ?? comp?.posX ?? comp?.positionX
	);
	const y = Number(
		(typeof comp?.getState_Y === 'function' ? comp.getState_Y() : undefined)
		?? (typeof comp?.getState_CenterY === 'function' ? comp.getState_CenterY() : undefined)
		?? comp?.y ?? comp?.Y ?? comp?.centerY ?? comp?.cy ?? comp?.posY ?? comp?.positionY
	);
	if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
	return { x, y };
}

function scoreOutlineCandidate(poly: Pt[], samplePts: Pt[]): { inside: number; areaAbs: number; bbox: number } {
	const areaAbs = Math.abs(polygonArea(poly));
	const bbox = bboxArea(poly);
	let inside = 0;
	for (const p of samplePts) {
		if (pointInPolygon(p, poly)) inside++;
	}
	return { inside, areaAbs, bbox };
}

function bboxOfPoints(pts: Pt[]): { minX: number; minY: number; maxX: number; maxY: number } | null {
	let minX = Infinity;
	let minY = Infinity;
	let maxX = -Infinity;
	let maxY = -Infinity;
	for (const p of pts) {
		if (!Number.isFinite(p.x) || !Number.isFinite(p.y)) continue;
		minX = Math.min(minX, p.x);
		minY = Math.min(minY, p.y);
		maxX = Math.max(maxX, p.x);
		maxY = Math.max(maxY, p.y);
	}
	if (!Number.isFinite(minX)) return null;
	return { minX, minY, maxX, maxY };
}

function bboxAreaFromBounds(b: { minX: number; minY: number; maxX: number; maxY: number }): number {
	return Math.max(0, b.maxX - b.minX) * Math.max(0, b.maxY - b.minY);
}

function pickBestOutline(candidates: Pt[][], samplePts: Pt[], opts?: { requireInside?: boolean }): Pt[] | null {
	if (!candidates.length) return null;
	if (!samplePts.length) {
		// Default: largest area
		return candidates.sort((a, b) => Math.abs(polygonArea(b)) - Math.abs(polygonArea(a)))[0];
	}

	// If the document contains unrelated outline loops far away (panel scraps / guide lines),
	// prefer candidates not wildly larger than the component spread.
	const sampleB = bboxOfPoints(samplePts);
	const sampleBBoxArea = sampleB ? bboxAreaFromBounds(sampleB) : 0;
	let filtered = candidates;
	if (sampleBBoxArea > 1) {
		const maxScale = 60; // generous: allows moderate panelization, avoids absurd triangles
		filtered = candidates.filter((c) => bboxArea(c) <= sampleBBoxArea * maxScale);
		if (!filtered.length) filtered = candidates;
	}

	let best: Pt[] | null = null;
	let bestInside = -1;
	let bestArea = Infinity;
	for (const c of filtered) {
		const s = scoreOutlineCandidate(c, samplePts);
		// Prefer containing more sample points; tie-breaker: smaller area (avoid huge triangles).
		if (s.inside > bestInside || (s.inside === bestInside && s.areaAbs < bestArea)) {
			bestInside = s.inside;
			bestArea = s.areaAbs;
			best = c;
		}
	}

	if (opts?.requireInside && bestInside <= 0) return null;
	// If we failed to contain any point (all 0), fall back to smallest reasonable area among top few,
	// but avoid degenerates: require at least 3 points and non-zero area.
	if (bestInside <= 0) {
		const valid = filtered.filter((c) => Math.abs(polygonArea(c)) > 1);
		if (valid.length) return valid.sort((a, b) => Math.abs(polygonArea(a)) - Math.abs(polygonArea(b)))[0];
	}
	return best;
}

function computeBounds(polys: Pt[][]): { minX: number; minY: number; maxX: number; maxY: number } | null {
	let minX = Infinity;
	let minY = Infinity;
	let maxX = -Infinity;
	let maxY = -Infinity;
	for (const poly of polys) {
		for (const p of poly) {
			if (!Number.isFinite(p.x) || !Number.isFinite(p.y)) continue;
			minX = Math.min(minX, p.x);
			minY = Math.min(minY, p.y);
			maxX = Math.max(maxX, p.x);
			maxY = Math.max(maxY, p.y);
		}
	}
	if (!Number.isFinite(minX)) return null;
	return { minX, minY, maxX, maxY };
}

function polyToSvgPath(poly: Pt[]): string {
	return poly.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x.toFixed(3)} ${p.y.toFixed(3)}`).join(' ') + ' Z';
}

function generateSvg(outline: Pt[], cutouts: Pt[][]): string {
	// SVG is y-down; our export uses y-down KLE coordinates.
	const b = computeBounds([outline, ...cutouts].filter((p) => p && p.length));
	if (!b) throw new Error('No geometry');
	const pad = 2;
	const minX = b.minX - pad;
	const minY = b.minY - pad;
	const maxX = b.maxX + pad;
	const maxY = b.maxY + pad;
	const w = maxX - minX;
	const h = maxY - minY;

	const outlinePath = polyToSvgPath(outline);
	const cutoutsPath = cutouts.length ? cutouts.map(polyToSvgPath).join(' ') : '';

	return `<?xml version="1.0" encoding="UTF-8"?>\n` +
		`<svg xmlns="http://www.w3.org/2000/svg" width="${w.toFixed(3)}mm" height="${h.toFixed(3)}mm" viewBox="${minX.toFixed(3)} ${minY.toFixed(3)} ${w.toFixed(3)} ${h.toFixed(3)}">\n` +
		`  <g fill="none" stroke="#000" stroke-width="0.2">\n` +
		`    <path d="${outlinePath}"/>\n` +
		(cutoutsPath ? `    <path d="${cutoutsPath}"/>\n` : '') +
		`  </g>\n` +
		`</svg>\n`;
}

function generateSvgMulti(named: Array<{ name: string; poly: Pt[] }>, cutouts: Pt[][]): string {
	const polys: Pt[][] = [];
	for (const n of named) polys.push(n.poly);
	for (const c of cutouts) polys.push(c);
	const b = computeBounds(polys.filter((p) => p && p.length));
	if (!b) throw new Error('No geometry');
	const pad = 2;
	const minX = b.minX - pad;
	const minY = b.minY - pad;
	const maxX = b.maxX + pad;
	const maxY = b.maxY + pad;
	const w = maxX - minX;
	const h = maxY - minY;

	const color = (name: string) => {
		if (name === 'CASE_OUTER') return '#00a000';
		if (name === 'PCB_CLEARANCE') return '#0000ff';
		return '#000';
	};

	let paths = '';
	for (const n of named) {
		paths += `    <path data-name="${n.name}" stroke="${color(n.name)}" d="${polyToSvgPath(n.poly)}"/>\n`;
	}
	const cutoutsPath = cutouts.length ? cutouts.map(polyToSvgPath).join(' ') : '';
	if (cutoutsPath) paths += `    <path data-name="KEY_HOLES" stroke="#b00020" d="${cutoutsPath}"/>\n`;

	return `<?xml version="1.0" encoding="UTF-8"?>\n` +
		`<svg xmlns="http://www.w3.org/2000/svg" width="${w.toFixed(3)}mm" height="${h.toFixed(3)}mm" viewBox="${minX.toFixed(3)} ${minY.toFixed(3)} ${w.toFixed(3)} ${h.toFixed(3)}">\n` +
		`  <g fill="none" stroke-width="0.2">\n` +
		paths +
		`  </g>\n` +
		`</svg>\n`;
}

function dxfLine(layer: string, x1: number, y1: number, x2: number, y2: number): string {
	const f = (n: number) => Number(n).toFixed(4);
	return `0\nLINE\n8\n${layer}\n10\n${f(x1)}\n20\n${f(y1)}\n11\n${f(x2)}\n21\n${f(y2)}\n`;
}

function polyToDxfLines(layer: string, poly: Pt[]): string {
	let out = '';
	for (let i = 0; i < poly.length; i++) {
		const a = poly[i];
		const b = poly[(i + 1) % poly.length];
		// DXF is y-up: flip y
		out += dxfLine(layer, a.x, -a.y, b.x, -b.y);
	}
	return out;
}

function generateDxf(outline: Pt[], cutouts: Pt[][]): string {
	let entities = '';
	entities += polyToDxfLines('OUTLINE', outline);
	for (const c of cutouts) entities += polyToDxfLines('CUTOUT', c);
	return `0\nSECTION\n2\nHEADER\n0\nENDSEC\n0\nSECTION\n2\nENTITIES\n${entities}0\nENDSEC\n0\nEOF\n`;
}

function generateDxfMulti(named: Array<{ layer: string; poly: Pt[] }>, cutouts: Pt[][]): string {
	let entities = '';
	for (const n of named) {
		entities += polyToDxfLines(n.layer, n.poly);
	}
	for (const c of cutouts) entities += polyToDxfLines('KEY_HOLES', c);
	return `0\nSECTION\n2\nHEADER\n0\nENDSEC\n0\nSECTION\n2\nENTITIES\n${entities}0\nENDSEC\n0\nEOF\n`;
}

function rectPointsUnits(cx: number, cy: number, size: number, rotDeg: number, applyRot: boolean): Pt[] {
	const h = size / 2;
	const pts: Pt[] = [
		{ x: cx - h, y: cy - h },
		{ x: cx + h, y: cy - h },
		{ x: cx + h, y: cy + h },
		{ x: cx - h, y: cy + h },
	];
	if (!applyRot || !rotDeg) return pts;
	const rad = (rotDeg * Math.PI) / 180;
	return pts.map((p) => {
		const dx = p.x - cx;
		const dy = p.y - cy;
		return {
			x: cx + (dx * Math.cos(rad) - dy * Math.sin(rad)),
			y: cy + (dx * Math.sin(rad) + dy * Math.cos(rad)),
		};
	});
}

function roundKey(p: Pt, step = 0.01): string {
	const rx = Math.round(p.x / step) * step;
	const ry = Math.round(p.y / step) * step;
	return `${rx.toFixed(4)},${ry.toFixed(4)}`;
}

type OutlineItem =
	| { kind: 'line'; a: Pt; b: Pt }
	| { kind: 'arc'; a: Pt; b: Pt; angleDeg: number };

function extractLineItem(line: any): OutlineItem | null {
	const sx = Number(
		(typeof line?.getState_StartX === 'function' ? line.getState_StartX() : undefined)
		?? line?.startX ?? line?.sx ?? line?.sX ?? line?.x1 ?? line?.xStart ?? line?.x0
	);
	const sy = Number(
		(typeof line?.getState_StartY === 'function' ? line.getState_StartY() : undefined)
		?? line?.startY ?? line?.sy ?? line?.sY ?? line?.y1 ?? line?.yStart ?? line?.y0
	);
	const ex = Number(
		(typeof line?.getState_EndX === 'function' ? line.getState_EndX() : undefined)
		?? line?.endX ?? line?.ex ?? line?.eX ?? line?.x2 ?? line?.xEnd ?? line?.x1
	);
	const ey = Number(
		(typeof line?.getState_EndY === 'function' ? line.getState_EndY() : undefined)
		?? line?.endY ?? line?.ey ?? line?.eY ?? line?.y2 ?? line?.yEnd ?? line?.y1
	);
	if (![sx, sy, ex, ey].every((n) => Number.isFinite(n))) return null;
	return { kind: 'line', a: { x: sx, y: sy }, b: { x: ex, y: ey } };
}

function extractArcItem(arc: any): OutlineItem | null {
	const sx = Number(
		(typeof arc?.getState_StartX === 'function' ? arc.getState_StartX() : undefined)
		?? arc?.startX ?? arc?.sx ?? arc?.sX
	);
	const sy = Number(
		(typeof arc?.getState_StartY === 'function' ? arc.getState_StartY() : undefined)
		?? arc?.startY ?? arc?.sy ?? arc?.sY
	);
	const ex = Number(
		(typeof arc?.getState_EndX === 'function' ? arc.getState_EndX() : undefined)
		?? arc?.endX ?? arc?.ex ?? arc?.eX
	);
	const ey = Number(
		(typeof arc?.getState_EndY === 'function' ? arc.getState_EndY() : undefined)
		?? arc?.endY ?? arc?.ey ?? arc?.eY
	);
	let angleDeg = Number(
		(typeof arc?.getState_ArcAngle === 'function' ? arc.getState_ArcAngle() : undefined)
		?? arc?.arcAngle ?? arc?.a ?? arc?.angle ?? arc?.angleDeg
	);
	if (![sx, sy, ex, ey].every((n) => Number.isFinite(n))) return null;
	if (!Number.isFinite(angleDeg)) {
		// Some arc primitives store start/end angles + center/radius; if so we can approximate using the sweep.
		const startAngle = Number(
			(typeof arc?.getState_StartAngle === 'function' ? arc.getState_StartAngle() : undefined)
			?? arc?.startAngle ?? arc?.sa
		);
		const endAngle = Number(
			(typeof arc?.getState_EndAngle === 'function' ? arc.getState_EndAngle() : undefined)
			?? arc?.endAngle ?? arc?.ea
		);
		if (Number.isFinite(startAngle) && Number.isFinite(endAngle)) {
			angleDeg = endAngle - startAngle;
		}
	}
	if (!Number.isFinite(angleDeg) || angleDeg === 0) return null;
	return { kind: 'arc', a: { x: sx, y: sy }, b: { x: ex, y: ey }, angleDeg };
}

function arcToPoints(item: Extract<OutlineItem, { kind: 'arc' }>, segmentsMin = 8): Pt[] {
	// Approximate arc with line segments (including start & end).
	const a = item.a;
	const b = item.b;
	const theta = (item.angleDeg * Math.PI) / 180;
	const absTheta = Math.abs(theta);
	const chord = Math.hypot(b.x - a.x, b.y - a.y);
	if (chord < 1e-6 || absTheta < 1e-6) return [a, b];

	// radius and offset to center from midpoint
	const radius = chord / (2 * Math.sin(absTheta / 2));
	const h = chord / (2 * Math.tan(absTheta / 2));
	const mx = (a.x + b.x) / 2;
	const my = (a.y + b.y) / 2;
	const vx = b.x - a.x;
	const vy = b.y - a.y;
	const vlen = Math.hypot(vx, vy) || 1;
	// Perp normal (y-up). Choose sign by theta.
	const nx = -vy / vlen;
	const ny = vx / vlen;
	const sign = theta > 0 ? 1 : -1;
	const cx = mx + nx * h * sign;
	const cy = my + ny * h * sign;

	const startAng = Math.atan2(a.y - cy, a.x - cx);
	const segs = Math.max(segmentsMin, Math.ceil((absTheta / (Math.PI / 2)) * 12)); // ~12 segments per 90deg
	const step = theta / segs;
	const pts: Pt[] = [];
	for (let i = 0; i <= segs; i++) {
		const ang = startAng + step * i;
		pts.push({ x: cx + radius * Math.cos(ang), y: cy + radius * Math.sin(ang) });
	}
	// Snap last point to exact end (reduce drift)
	pts[pts.length - 1] = { x: b.x, y: b.y };
	return pts;
}

function buildClosedPathFromItems(items: OutlineItem[], tolStep = 0.01): Pt[] | null {
	// Greedy walk. Works well for typical board outlines where each vertex has degree 2.
	if (!items.length) return null;
	const unused = new Set<number>(items.map((_, i) => i));

	const endpointMap = new Map<string, Array<{ idx: number; end: 'a' | 'b' }>>();
	const addEnd = (idx: number, end: 'a' | 'b', p: Pt) => {
		const k = roundKey(p, tolStep);
		const arr = endpointMap.get(k) || [];
		arr.push({ idx, end });
		endpointMap.set(k, arr);
	};
	for (let i = 0; i < items.length; i++) {
		addEnd(i, 'a', (items[i] as any).a);
		addEnd(i, 'b', (items[i] as any).b);
	}

	const pickStart = () => {
		// Prefer a point with odd degree (if broken), else just first.
		for (const [k, arr] of endpointMap.entries()) {
			if ((arr.length % 2) === 1) return k;
		}
		return roundKey((items[0] as any).a, tolStep);
	};

	const startKey = pickStart();
	const startCandidates = endpointMap.get(startKey) || [];
	if (!startCandidates.length) return null;

	let currentKey = startKey;
	let currentPoint: Pt = (items[startCandidates[0].idx] as any)[startCandidates[0].end];
	let startPoint: Pt = currentPoint;
	let prevDir: Pt | null = null;
	const path: Pt[] = [startPoint];

	const chooseNext = (cands: Array<{ idx: number; end: 'a' | 'b' }>) => {
		const available = cands.filter((c) => unused.has(c.idx));
		if (available.length <= 1) return available[0] || null;
		if (!prevDir) return available[0];

		let best: any = null;
		let bestScore = -Infinity;
		for (const c of available) {
			const it: any = items[c.idx];
			const pThis = it[c.end] as Pt;
			const pOther = it[c.end === 'a' ? 'b' : 'a'] as Pt;
			const dx = pOther.x - pThis.x;
			const dy = pOther.y - pThis.y;
			const len = Math.hypot(dx, dy) || 1;
			const dir = { x: dx / len, y: dy / len };
			const score = dir.x * prevDir.x + dir.y * prevDir.y; // prefer straight
			if (score > bestScore) { bestScore = score; best = c; }
		}
		return best;
	};

	for (let steps = 0; steps < items.length + 10; steps++) {
		const cands = endpointMap.get(currentKey) || [];
		const next = chooseNext(cands);
		if (!next) break;

		const idx = next.idx;
		const it = items[idx];
		unused.delete(idx);

		const thisPt = next.end === 'a' ? (it as any).a as Pt : (it as any).b as Pt;
		const otherPt = next.end === 'a' ? (it as any).b as Pt : (it as any).a as Pt;

		// Update prevDir
		const ddx = otherPt.x - thisPt.x;
		const ddy = otherPt.y - thisPt.y;
		const dlen = Math.hypot(ddx, ddy) || 1;
		prevDir = { x: ddx / dlen, y: ddy / dlen };

		if (it.kind === 'arc') {
			const pts = arcToPoints(it, 8);
			if (next.end === 'b') pts.reverse();
			// pts includes start; avoid duplicating current point
			for (let i = 1; i < pts.length; i++) path.push(pts[i]);
			currentPoint = pts[pts.length - 1];
		}
		else {
			path.push(otherPt);
			currentPoint = otherPt;
		}
		currentKey = roundKey(currentPoint, tolStep);

		// Closed?
		if (currentKey === roundKey(startPoint, tolStep)) {
			return path;
		}
	}

	return null;
}

function extractCyclesFromItems(items: OutlineItem[], tolStep = 0.01): Pt[][] {
	if (!items.length) return [];

	const endpointMap = new Map<string, Array<{ idx: number; end: 'a' | 'b' }>>();
	const addEnd = (idx: number, end: 'a' | 'b', p: Pt) => {
		const k = roundKey(p, tolStep);
		const arr = endpointMap.get(k) || [];
		arr.push({ idx, end });
		endpointMap.set(k, arr);
	};
	for (let i = 0; i < items.length; i++) {
		addEnd(i, 'a', (items[i] as any).a);
		addEnd(i, 'b', (items[i] as any).b);
	}

	const degree = (k: string) => (endpointMap.get(k) || []).length;
	const goodEdge = (i: number) => {
		const aKey = roundKey((items[i] as any).a, tolStep);
		const bKey = roundKey((items[i] as any).b, tolStep);
		return degree(aKey) >= 2 && degree(bKey) >= 2;
	};

	const unused = new Set<number>();
	for (let i = 0; i < items.length; i++) {
		if (goodEdge(i)) unused.add(i);
	}

	const cycles: Pt[][] = [];

	const chooseNext = (currentKey: string, prevIdx: number | null, prevDir: Pt | null) => {
		const cands = (endpointMap.get(currentKey) || []).filter((c) => unused.has(c.idx) && c.idx !== prevIdx);
		if (!cands.length) return null;
		if (cands.length === 1 || !prevDir) return cands[0];

		let best: any = null;
		let bestScore = -Infinity;
		for (const c of cands) {
			const it: any = items[c.idx];
			const pThis = it[c.end] as Pt;
			const pOther = it[c.end === 'a' ? 'b' : 'a'] as Pt;
			const dx = pOther.x - pThis.x;
			const dy = pOther.y - pThis.y;
			const len = Math.hypot(dx, dy) || 1;
			const dir = { x: dx / len, y: dy / len };
			const score = dir.x * prevDir.x + dir.y * prevDir.y;
			if (score > bestScore) { bestScore = score; best = c; }
		}
		return best;
	};

	const walkFromEdge = (startIdx: number) => {
		const it0: any = items[startIdx];
		const startPoint = it0.a as Pt;
		const startKey = roundKey(startPoint, tolStep);
		let currentKey = startKey;
		let prevIdx: number | null = null;
		let prevDir: Pt | null = null;
		const path: Pt[] = [startPoint];

		for (let steps = 0; steps < items.length + 10; steps++) {
			const next = chooseNext(currentKey, prevIdx, prevDir);
			if (!next) return null;
			const idx = next.idx;
			const it: any = items[idx];
			unused.delete(idx);
			prevIdx = idx;

			const thisPt = next.end === 'a' ? it.a as Pt : it.b as Pt;
			const otherPt = next.end === 'a' ? it.b as Pt : it.a as Pt;

			const ddx = otherPt.x - thisPt.x;
			const ddy = otherPt.y - thisPt.y;
			const dlen = Math.hypot(ddx, ddy) || 1;
			prevDir = { x: ddx / dlen, y: ddy / dlen };

			if (it.kind === 'arc') {
				const pts = arcToPoints(it, 8);
				if (next.end === 'b') pts.reverse();
				for (let i = 1; i < pts.length; i++) path.push(pts[i]);
				currentKey = roundKey(pts[pts.length - 1], tolStep);
			}
			else {
				path.push(otherPt);
				currentKey = roundKey(otherPt, tolStep);
			}

			if (currentKey === startKey && path.length >= 4) {
				return path;
			}
		}
		return null;
	};

	for (const idx of [...unused]) {
		if (!unused.has(idx)) continue;
		const cycle = walkFromEdge(idx);
		if (cycle && cycle.length >= 3) cycles.push(cycle);
	}

	return cycles;
}

function offsetPolygonClipper(points: Pt[], delta: number): Pt[] | null {
	// Uses clipper-lib to robustly offset even concave polygons.
	// `delta` is in the same units as `points` (mil in our usage).
	if (!points || points.length < 3) return null;
	const scale = 1000; // increase precision; Clipper uses integer coords
	const subj = points.map((p) => ({ X: Math.round(p.x * scale), Y: Math.round(p.y * scale) }));

	const doOffset = (d: number) => {
		const co = new (ClipperLib as any).ClipperOffset(2, 0.25 * scale);
		co.AddPath(subj, (ClipperLib as any).JoinType.jtMiter, (ClipperLib as any).EndType.etClosedPolygon);
		const solution: any[] = [];
		co.Execute(solution, d * scale);
		const paths = Array.isArray(solution) ? solution : [];
		if (!paths.length) return null;
		// Choose largest by area.
		let best: any[] | null = null;
		let bestAbsArea = -Infinity;
		for (const path of paths) {
			if (!Array.isArray(path) || path.length < 3) continue;
			const pts = path.map((p: any) => ({ x: p.X / scale, y: p.Y / scale }));
			const aa = Math.abs(polygonArea(pts));
			if (aa > bestAbsArea) {
				bestAbsArea = aa;
				best = path;
			}
		}
		if (!best) return null;
		return best.map((p: any) => ({ x: p.X / scale, y: p.Y / scale }));
	};

	let out = doOffset(delta);
	if (!out) return null;
	// Sanity: expansion should generally increase bbox area. If it shrank, flip delta.
	if (bboxArea(out) < bboxArea(points)) {
		const flipped = doOffset(-delta);
		if (flipped && bboxArea(flipped) > bboxArea(out)) out = flipped;
	}

	// Remove duplicate closing point if present.
	if (out.length >= 2) {
		const a = out[0];
		const b = out[out.length - 1];
		if (Math.hypot(a.x - b.x, a.y - b.y) < 1e-6) out = out.slice(0, -1);
	}
	return out.length >= 3 ? out : null;
}

function rectPointsMil(cx: number, cy: number, sizeMil: number, rotDeg: number, applyRot: boolean): Pt[] {
	const h = sizeMil / 2;
	const pts: Pt[] = [
		{ x: cx - h, y: cy - h },
		{ x: cx + h, y: cy - h },
		{ x: cx + h, y: cy + h },
		{ x: cx - h, y: cy + h },
	];
	if (!applyRot || !rotDeg) return pts;
	const rad = (rotDeg * Math.PI) / 180;
	return pts.map((p) => {
		const dx = p.x - cx;
		const dy = p.y - cy;
		return {
			x: cx + (dx * Math.cos(rad) - dy * Math.sin(rad)),
			y: cy + (dx * Math.sin(rad) + dy * Math.cos(rad)),
		};
	});
}

function convexHull(points: Pt[]): Pt[] {
	// Monotonic chain, returns hull in CCW (y-up) without repeating first point.
	const pts = points
		.filter((p) => Number.isFinite(p.x) && Number.isFinite(p.y))
		.map((p) => ({ x: p.x, y: p.y }));
	if (pts.length <= 1) return pts;

	pts.sort((a, b) => (a.x - b.x) || (a.y - b.y));
	const cross = (o: Pt, a: Pt, b: Pt) => (a.x - o.x) * (b.y - o.y) - (a.y - o.y) * (b.x - o.x);

	const lower: Pt[] = [];
	for (const p of pts) {
		while (lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], p) <= 0) lower.pop();
		lower.push(p);
	}
	const upper: Pt[] = [];
	for (let i = pts.length - 1; i >= 0; i--) {
		const p = pts[i];
		while (upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], p) <= 0) upper.pop();
		upper.push(p);
	}
	upper.pop();
	lower.pop();
	return lower.concat(upper);
}

function unionPolygonsClipper(polys: Pt[][]): Pt[][] {
	// Boolean union of polygons. Returns one or more resulting closed paths.
	const scale = 1000;
	const toPath = (poly: Pt[]) => poly.map((p) => ({ X: Math.round(p.x * scale), Y: Math.round(p.y * scale) }));
	const subj = polys.filter((p) => p && p.length >= 3).map(toPath);
	if (!subj.length) return [];

	const clipper = new (ClipperLib as any).Clipper();
	clipper.AddPaths(subj, (ClipperLib as any).PolyType.ptSubject, true);
	const solution: any[] = [];
	clipper.Execute(
		(ClipperLib as any).ClipType.ctUnion,
		solution,
		(ClipperLib as any).PolyFillType.pftNonZero,
		(ClipperLib as any).PolyFillType.pftNonZero,
	);

	const paths = Array.isArray(solution) ? solution : [];
	// Optional cleanup/simplify
	let simplified = paths;
	try {
		if ((ClipperLib as any).Clipper?.SimplifyPolygons) {
			simplified = (ClipperLib as any).Clipper.SimplifyPolygons(paths, (ClipperLib as any).PolyFillType.pftNonZero);
		}
	}
	catch {}

	const out: Pt[][] = [];
	for (const path of simplified) {
		if (!Array.isArray(path) || path.length < 3) continue;
		out.push(path.map((p: any) => ({ x: p.X / scale, y: p.Y / scale })));
	}
	return out;
}

function toMmFromMil(poly: Pt[]): Pt[] {
	const k = 25.4 / 1000;
	return poly.map((p) => ({ x: p.x * k, y: p.y * k }));
}

function flipY(poly: Pt[]): Pt[] {
	return poly.map((p) => ({ x: p.x, y: -p.y }));
}

function buildKeyHolesFromKleMm(rawJson: string, pitchMm: number, sizeMm: number, applyRot: boolean): Pt[][] {
	const kle = parseKleText(rawJson);
	const layout = parseKLE(kle, 0, 0);
	const holes: Pt[][] = [];
	for (const k of layout) {
		const cx = Number(k?.swX) * pitchMm;
		const cy = -Number(k?.swY) * pitchMm; // PCB y-up
		const rot = -Number(k?.rot || 0);
		if (!Number.isFinite(cx) || !Number.isFinite(cy)) continue;
		holes.push(rectPointsUnits(cx, cy, sizeMm, rot, applyRot));
	}
	return holes;
}

function getSamplePointsFromKleMil(rawJson: string, pitchMm: number): Pt[] {
	try {
		const kle = parseKleText(rawJson);
		const milPerU = (pitchMm / 25.4) * 1000;
		const layout = parseKLE(kle, 0, 0);
		const pts: Pt[] = [];
		for (const k of layout) {
			const x = Number(k?.swX) * milPerU;
			const y = -Number(k?.swY) * milPerU;
			if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
			pts.push({ x, y });
		}
		return pts;
	}
	catch {
		return [];
	}
}

async function getSamplePointsFromComponents(): Promise<Pt[]> {
	try {
		const compsRaw = await (eda as any)?.pcb_PrimitiveComponent?.getAll?.();
		const comps = asArray<any>(compsRaw);
		if (!comps.length) return [];

		const swd = comps.filter((c) => {
			const d = String(c?.designator ?? c?.getState_Designator?.() ?? '');
			return /^SW\d+$/i.test(d) || /^D\d+$/i.test(d);
		});
		let sample = swd.map(extractComponentPoint).filter((p): p is Pt => Boolean(p));
		if (sample.length >= 3) return sample;

		sample = comps.map(extractComponentPoint).filter((p): p is Pt => Boolean(p));
		return sample;
	}
	catch {
		return [];
	}
}

async function getSwitchFootprintCentersMil(): Promise<Pt[]> {
	try {
		const compsRaw = await (eda as any)?.pcb_PrimitiveComponent?.getAll?.();
		const comps = asArray<any>(compsRaw);
		if (!comps.length) return [];

		const sw = comps.filter((c) => {
			const d = String(c?.designator ?? c?.getState_Designator?.() ?? '');
			return /^SW\d+$/i.test(d);
		});

		return sw.map(extractComponentPoint).filter((p): p is Pt => Boolean(p));
	}
	catch {
		return [];
	}
}

async function buildOutlineFromFootprintsMm(mode: 'plateHull' | 'rectUnion', rectSizeMm: number, marginMm: number): Promise<Pt[]> {
	const swMil = await getSwitchFootprintCentersMil();
	const sampleMil = swMil.length ? swMil : await getSamplePointsFromComponents();
	if (sampleMil.length < 1) {
		throw new Error('No component positions found. Place/import SW footprints first, or use BoardOutline as outline source.');
	}

	const mmPerMil = 25.4 / 1000;
	const rects: Pt[][] = [];
	for (const p of sampleMil) {
		const cx = p.x * mmPerMil;
		const cy = p.y * mmPerMil;
		if (!Number.isFinite(cx) || !Number.isFinite(cy)) continue;
		const base = rectPointsUnits(cx, cy, rectSizeMm, 0, false);
		if (marginMm > 0) {
			const expanded = offsetPolygonClipper(base, marginMm);
			rects.push(expanded ?? base);
		}
		else {
			rects.push(base);
		}
	}
	if (!rects.length) throw new Error('Failed to build footprint rectangles');

	if (mode === 'plateHull') {
		const hull = convexHull(rects.flat());
		// 'marginMm' already applied per-rect; for hull mode, also apply it to hull so sparse layouts still connect.
		const outline = offsetPolygonClipper(hull, marginMm) ?? hull;
		return outline;
	}

	const union = unionPolygonsClipper(rects);
	const picked = union.sort((a, b) => Math.abs(polygonArea(b)) - Math.abs(polygonArea(a)))[0];
	if (!picked) throw new Error('Union failed');
	return picked;
}

async function exportSwitchPlate(payload: PlateExportPayload) {
	const rawJson = payload?.json;
	if (!rawJson) throw new Error('Missing KLE JSON');

	const pitchMm = Number(payload.pitch);
	if (!Number.isFinite(pitchMm) || pitchMm <= 0) throw new Error('Invalid pitch');

	const fmt = (payload.format === 'dxf') ? 'dxf' : 'svg';
	const marginMm = Number.isFinite(Number(payload.marginMm)) ? Number(payload.marginMm) : 5.0;
	const cutoutMm = Number.isFinite(Number(payload.cutoutSizeMm)) ? Number(payload.cutoutSizeMm) : 14.0;
	const applyRot = payload.applyKeyRotation !== false;
	const outlineMode = (payload.outlineMode === 'rectUnion') ? 'rectUnion' : 'hull';
	const rectUnionMarginMm = Number.isFinite(Number(payload.rectUnionMarginMm)) ? Number(payload.rectUnionMarginMm) : 2.0;

	const kle = parseKleText(rawJson);
	const layout = parseKLE(kle, 0, 0);

	const cutouts: Pt[][] = [];
	for (const k of layout) {
		const cx = Number(k?.swX) * pitchMm;
		const cy = Number(k?.swY) * pitchMm;
		const rot = Number(k?.rot || 0);
		if (!Number.isFinite(cx) || !Number.isFinite(cy)) continue;
		// Export uses y-down and parseKLE returns y-down, so use +cy.
		cutouts.push(rectPointsUnits(cx, cy, cutoutMm, rot, applyRot));
	}
	if (!cutouts.length) throw new Error('No keys found');

	const allPts = cutouts.flat();
	let outline: Pt[] | null = null;
	if (outlineMode === 'hull') {
		const hull = convexHull(allPts);
		outline = offsetPolygonClipper(hull, marginMm) ?? hull;
	}
	else {
		const rectsExpanded = cutouts.map((poly) => {
			// Expand each cutout by rectUnionMarginMm before union (approx by offsetting the rectangle polygon).
			const expanded = offsetPolygonClipper(poly, rectUnionMarginMm);
			return expanded ?? poly;
		});
		const union = unionPolygonsClipper(rectsExpanded);
		const picked = union.sort((a, b) => Math.abs(polygonArea(b)) - Math.abs(polygonArea(a)))[0] ?? null;
		if (!picked) throw new Error('Union failed');
		outline = offsetPolygonClipper(picked, marginMm) ?? picked;
	}
	if (!outline) throw new Error('Outline generation failed');

	const base = String(payload?.fileName || 'kle').replace(/\.[^.]+$/, '');
	const filename = fmt === 'dxf' ? `switch-plate_${base}.dxf` : `switch-plate_${base}.svg`;
	const text = fmt === 'dxf' ? generateDxf(outline, cutouts) : generateSvg(outline, cutouts);
	return { ok: true, filename, format: fmt, text };
}

async function extractBoardOutlineMm(outlineSource: 'board' | 'selected', samplePtsMil?: Pt[]): Promise<Pt[]> {
	const LAYER_BOARD_OUTLINE = 11;
	const polyApi: any = (eda as any)?.pcb_PrimitivePolyline;
	let baseOutline: Pt[] | null = null;
	const samplePts = (samplePtsMil && samplePtsMil.length >= 3) ? samplePtsMil : await getSamplePointsFromComponents();
	const requireInside = samplePts.length >= 3;

	const useSelected = outlineSource === 'selected';
	if (useSelected && (eda as any)?.pcb_SelectControl?.getAllSelectedPrimitives_PrimitiveId) {
		const selectedIds = asArray<any>(await (eda as any).pcb_SelectControl.getAllSelectedPrimitives_PrimitiveId());
		const selectedSet = new Set<string>(selectedIds.filter((id: any) => typeof id === 'string').map((id: string) => String(id)));
		if (!selectedSet.size) {
			throw new Error('No primitives selected. Select the correct board outline and retry.');
		}
		const isSelected = (p: any) => {
			const pid =
				(typeof p?.getState_PrimitiveId === 'function' ? p.getState_PrimitiveId() : undefined)
				?? p?.primitiveId ?? p?.id;
			return pid && selectedSet.has(String(pid));
		};

		if (polyApi?.getAll) {
			const allPolys = asArray<any>(await polyApi.getAll());
			const selPolys = allPolys.filter(isSelected);
			const candidates = selPolys
				.map((p) => extractPolylinePoints(p))
				.filter((pts) => pts && pts.length >= 3) as Pt[][];
			baseOutline = pickBestOutline(candidates, samplePts, { requireInside });
		}
		if (!baseOutline) {
			const lineApi: any = (eda as any)?.pcb_PrimitiveLine;
			const arcApi: any = (eda as any)?.pcb_PrimitiveArc;
			const lines = lineApi?.getAll ? asArray<any>(await lineApi.getAll()) : [];
			const arcs = arcApi?.getAll ? asArray<any>(await arcApi.getAll()) : [];
			const items: OutlineItem[] = [];
			for (const l of lines) { if (isSelected(l)) { const it = extractLineItem(l); if (it) items.push(it); } }
			for (const a of arcs) { if (isSelected(a)) { const it = extractArcItem(a); if (it) items.push(it); } }
			const cycles = extractCyclesFromItems(items, 0.01);
			baseOutline = pickBestOutline(cycles, samplePts, { requireInside });
		}

		if (!baseOutline) {
			throw new Error('Selected primitives do not form a closed outline containing the design.');
		}
	}

	if (!baseOutline && polyApi?.getAll) {
		const allPolys = asArray<any>(await polyApi.getAll());
		const candidates = allPolys
			.filter((p) => getLayerId(p) === LAYER_BOARD_OUTLINE)
			.map((p) => extractPolylinePoints(p))
			.filter((pts) => pts && pts.length >= 3) as Pt[][];
		baseOutline = pickBestOutline(candidates, samplePts, { requireInside });
	}

	if (!baseOutline) {
		const lineApi: any = (eda as any)?.pcb_PrimitiveLine;
		const arcApi: any = (eda as any)?.pcb_PrimitiveArc;
		const lines = lineApi?.getAll ? asArray<any>(await lineApi.getAll()) : [];
		const arcs = arcApi?.getAll ? asArray<any>(await arcApi.getAll()) : [];
		const items: OutlineItem[] = [];
		for (const l of lines) { if (getLayerId(l) === LAYER_BOARD_OUTLINE) { const it = extractLineItem(l); if (it) items.push(it); } }
		for (const a of arcs) { if (getLayerId(a) === LAYER_BOARD_OUTLINE) { const it = extractArcItem(a); if (it) items.push(it); } }
		const cycles = extractCyclesFromItems(items, 0.01);
		baseOutline = pickBestOutline(cycles, samplePts, { requireInside });
	}

	if (!baseOutline) {
		throw new Error('BOARD_OUTLINE not found or does not contain the design. Use "selected outline source" and select the correct outline.');
	}

	// If the outline is still wildly larger than components, it's likely a stray loop.
	try {
		const sb = bboxOfPoints(samplePts);
		if (sb) {
			const sArea = bboxAreaFromBounds(sb);
			const oArea = bboxArea(baseOutline);
			if (sArea > 1 && oArea > sArea * 80) {
				throw new Error('BOARD_OUTLINE seems unrelated (too large). Use "selected outline source" and select the correct outline.');
			}
		}
	}
	catch (e) {
		if (e instanceof Error) throw e;
	}
	return toMmFromMil(baseOutline);
}

async function exportCase(payload: CaseExportPayload) {
	const rawJson = payload?.json;
	if (!rawJson) throw new Error('Missing KLE JSON');

	const pitchMm = Number(payload.pitch);
	if (!Number.isFinite(pitchMm) || pitchMm <= 0) throw new Error('Invalid pitch');

	const fmt = (payload.format === 'dxf') ? 'dxf' : 'svg';
	const clearanceMm = Number.isFinite(Number(payload.boardClearanceMm)) ? Number(payload.boardClearanceMm) : 0.2;
	const wallMm = Number.isFinite(Number(payload.wallThicknessMm)) ? Number(payload.wallThicknessMm) : 2.0;
	const includeHoles = payload.includeKeyHoles !== false;
	const holeSizeMm = Number.isFinite(Number(payload.keyHoleSizeMm)) ? Number(payload.keyHoleSizeMm) : 14.0;
	const applyRot = payload.applyKeyRotation !== false;
	const outlineSourceRaw = String(payload.outlineSource || 'board');
	const outlineSource =
		outlineSourceRaw === 'selected' ? 'selected'
			: (outlineSourceRaw === 'plateHull' ? 'plateHull'
				: (outlineSourceRaw === 'rectUnion' ? 'rectUnion' : 'board'));

	let board: Pt[] = [];
	if (outlineSource === 'plateHull' || outlineSource === 'rectUnion') {
		const rectUnionMarginMm = Number.isFinite(Number(payload.rectUnionMarginMm)) ? Number(payload.rectUnionMarginMm) : 2.0;
		const rectSizeMm = Number.isFinite(Number(payload.outlineRectSizeMm)) ? Number(payload.outlineRectSizeMm) : pitchMm;
		board = await buildOutlineFromFootprintsMm(outlineSource, rectSizeMm, rectUnionMarginMm);
	}
	else {
		// Board-outline selection must be anchored to the actual design.
		// Component positions are not always readable in some EasyEDA builds/environments,
		// so prefer KLE-derived sample points to pick the correct outline loop.
		const kleSamplePtsMil = getSamplePointsFromKleMil(rawJson, pitchMm);
		board = await extractBoardOutlineMm(outlineSource, kleSamplePtsMil);
	}
	const pcbClear = offsetPolygonClipper(board, clearanceMm) ?? board;
	const caseOuter = offsetPolygonClipper(board, clearanceMm + wallMm) ?? pcbClear;

	let holes: Pt[][] = [];
	if (includeHoles) {
		holes = buildKeyHolesFromKleMm(rawJson, pitchMm, holeSizeMm, applyRot);
	}

	// For SVG, flip Y to y-down for easier viewing.
	const named = [
		{ name: 'PCB_CLEARANCE', layer: 'PCB_CLEARANCE', poly: pcbClear },
		{ name: 'CASE_OUTER', layer: 'CASE_OUTER', poly: caseOuter },
	];
	const base = String(payload?.fileName || 'kle').replace(/\.[^.]+$/, '');

	if (fmt === 'svg') {
		const svg = generateSvgMulti(
			named.map((n) => ({ name: n.name, poly: flipY(n.poly) })),
			holes.map(flipY),
		);
		return { ok: true, filename: `case_${base}.svg`, format: 'svg', text: svg };
	}
	else {
		const dxf = generateDxfMulti(
			named.map((n) => ({ layer: n.layer, poly: n.poly })),
			holes,
		);
		return { ok: true, filename: `case_${base}.dxf`, format: 'dxf', text: dxf };
	}
}

/*
 * 3D Shell generation (experimental) was removed from the extension UI/API.
 * We keep this block commented out for now to avoid accidental use and to
 * keep the build stable across EasyEDA variants.

async function deletePolylinesOnLayers(layers: number[]) {
	const api: any = (eda as any)?.pcb_PrimitivePolyline;
	if (!api?.getAll || !api?.delete) return;
	const all = asArray<any>(await api.getAll());
	const toDelete = all.filter((p) => layers.includes(Number(p?.layerId ?? p?.layer ?? p?.l))).map((p) => p?.primitiveId ?? p?.id ?? p?.getState_PrimitiveId?.());
	const ids = toDelete.filter((id: any) => id != null).map((id: any) => String(id));
	if (!ids.length) return;
	try {
		await api.delete(ids);
	}
	catch {
		// Some APIs expect [id] per call.
		for (const id of ids) {
			try { await api.delete([id]); } catch {}
		}
	}
}

async function deleteLinesOnLayers(layers: number[]) {
	const api: any = (eda as any)?.pcb_PrimitiveLine;
	if (!api?.getAll || !api?.delete) return;
	const all = asArray<any>(await api.getAll());
	const toDelete = all
		.filter((p) => layers.includes(getLayerId(p) ?? -1))
		.map((p) => p?.primitiveId ?? p?.id ?? p?.getState_PrimitiveId?.());
	const ids = toDelete.filter((id: any) => id != null).map((id: any) => String(id));
	if (!ids.length) return;
	try {
		await api.delete(ids);
	}
	catch {
		for (const id of ids) {
			try { await api.delete([id]); } catch {}
		}
	}
}

async function createClosedPolyline(layerId: number, pts: Pt[], label?: string): Promise<string | null> {
	const api: any = (eda as any)?.pcb_PrimitivePolyline;
	if (!api?.create) throw new Error('pcb_PrimitivePolyline.create not available');
	const ptsNorm = normalizePoints(pts, true);
	if (!ptsNorm || ptsNorm.length < 3) throw new Error('polyline requires >= 3 points');

	const netStr = '';
	const netNull = null;
	const lineWidth = 10; // mil (visual only)
	const primitiveLock = false;

	const ptsXY = ptsNorm.map((p) => ({ x: p.x, y: p.y }));
	const ptsXYCap = ptsNorm.map((p) => ({ X: p.x, Y: p.y }));
	const ptsStrSpace = ptsNorm.map((p) => `${p.x} ${p.y}`).join(' ');
	const ptsStrComma = ptsNorm.map((p) => `${p.x},${p.y}`).join(' ');

	const mathApi: any = (eda as any)?.pcb_MathPolygon;
	const polygon = await createPolygonRobust(mathApi, ptsNorm);

	const deleteById = async (id: string) => {
		try {
			if (api.delete) await api.delete([id]);
		}
		catch {}
	};

	const findCreated = async (id: string) => {
		try {
			if (!api.getAll) return null;
			const all = asArray<any>(await api.getAll());
			for (const p of all) {
				const pid =
					(typeof p?.getState_PrimitiveId === 'function' ? p.getState_PrimitiveId() : undefined)
					?? p?.primitiveId ?? p?.id;
				if (pid && String(pid) === id) return p;
			}
		}
		catch {}
		return null;
	};

	const tryMoveToLayer = async (id: string) => {
		try {
			if (!api.modify) return false;
			// Try both field names seen in other APIs.
			await api.modify(id, { layerId });
		}
		catch {
			try {
				await api.modify(id, { layer: layerId });
			}
			catch {
				return false;
			}
		}
		const created = await findCreated(id);
		return created ? getLayerId(created) === layerId : false;
	};

	const variants: Array<{ desc: string; args: any[] }> = [
		// Common variants seen in different EasyEDA builds (points array / string)
		{ desc: 'netStr,layer,pointsXY,width', args: [netStr, layerId as any, ptsXY, lineWidth] },
		{ desc: 'netNull,layer,pointsXY,width', args: [netNull, layerId as any, ptsXY, lineWidth] },
		{ desc: 'netStr,layer,pointsXY,width,lock', args: [netStr, layerId as any, ptsXY, lineWidth, primitiveLock] },
		{ desc: 'netNull,layer,pointsXY,width,lock', args: [netNull, layerId as any, ptsXY, lineWidth, primitiveLock] },
		{ desc: 'netStr,layer,pointsXY,width,lock,closed', args: [netStr, layerId as any, ptsXY, lineWidth, primitiveLock, true] },
		{ desc: 'netNull,layer,pointsXY,width,lock,closed', args: [netNull, layerId as any, ptsXY, lineWidth, primitiveLock, true] },
		{ desc: 'layer,pointsXY,width', args: [layerId as any, ptsXY, lineWidth] },
		{ desc: 'netStr,layer,pointsXYCap,width', args: [netStr, layerId as any, ptsXYCap, lineWidth] },
		{ desc: 'netNull,layer,pointsXYCap,width', args: [netNull, layerId as any, ptsXYCap, lineWidth] },
		{ desc: 'netStr,layer,pointsStrSpace,width', args: [netStr, layerId as any, ptsStrSpace, lineWidth] },
		{ desc: 'netNull,layer,pointsStrSpace,width', args: [netNull, layerId as any, ptsStrSpace, lineWidth] },
		{ desc: 'netStr,layer,pointsStrComma,width', args: [netStr, layerId as any, ptsStrComma, lineWidth] },
		{ desc: 'netNull,layer,pointsStrComma,width', args: [netNull, layerId as any, ptsStrComma, lineWidth] },
		{ desc: 'options(points)', args: [{ net: netStr, layerId, layer: layerId, points: ptsXY, lineWidth, width: lineWidth, lock: primitiveLock, primitiveLock, closed: true, name: label }] },
	];

	if (polygon) {
		variants.unshift(
			{ desc: 'netStr,layer,polygon,width,lock', args: [netStr, layerId as any, polygon, lineWidth, primitiveLock] },
			{ desc: 'netNull,layer,polygon,width,lock', args: [netNull, layerId as any, polygon, lineWidth, primitiveLock] },
			{ desc: 'netStr,polygon,layer,width,lock', args: [netStr, polygon, layerId as any, lineWidth, primitiveLock] },
			{ desc: 'netNull,polygon,layer,width,lock', args: [netNull, polygon, layerId as any, lineWidth, primitiveLock] },
			{ desc: 'layer,polygon,width,lock', args: [layerId as any, polygon, lineWidth, primitiveLock] },
			{ desc: 'options(polygon)', args: [{ net: netStr, layerId, layer: layerId, polygon, lineWidth, width: lineWidth, lock: primitiveLock, primitiveLock, name: label }] },
		);
	}

	let lastErr: any = null;
	let lastDesc = '';
	for (const v of variants) {
		try {
			lastDesc = v.desc;
			const res = await api.create.apply(api, v.args);
			const id = getPrimitiveIdFromUnknown(res);
			if (!id) {
				// Can't validate without an id; try next signature.
				continue;
			}
			const created = await findCreated(id);
			const actualLayer = created ? getLayerId(created) : null;
			if (actualLayer === layerId) return id;

			// Some builds ignore the layer argument; attempt to move it after creation.
			if (await tryMoveToLayer(id)) return id;

			// Wrong layer: delete and try next signature.
			await deleteById(id);
		}
		catch (e: any) {
			lastErr = e;
		}
	}

	throw new Error(
		`${label ? `${label}: ` : ''}pcb_PrimitivePolyline.create failed or created on wrong layer.\n` +
		`Tried ${variants.length} signatures. fn.length=${Number(api.create?.length)}. last=${lastDesc}.\n` +
		`Last error: ${lastErr?.message ?? String(lastErr)}`
	);
}

async function createClosedLoopAsLines(layerId: number, pts: Pt[], label?: string): Promise<string[]> {
	const api: any = (eda as any)?.pcb_PrimitiveLine;
	const fn: any = api?.create;
	if (!fn) throw new Error('pcb_PrimitiveLine.create not available');
	if (!pts || pts.length < 2) return [];

	const width = 10;
	const created: string[] = [];
	for (let i = 0; i < pts.length; i++) {
		const a = pts[i];
		const b = pts[(i + 1) % pts.length];
		if (!Number.isFinite(a.x) || !Number.isFinite(a.y) || !Number.isFinite(b.x) || !Number.isFinite(b.y)) continue;
		try {
			const res = await fn.call(api, '', layerId as any, a.x, a.y, b.x, b.y, width);
			const id = getPrimitiveIdFromUnknown(res);
			if (id) created.push(id);
		}
		catch (e: any) {
			throw new Error(`${label ? `${label}: ` : ''}pcb_PrimitiveLine.create failed: ${e?.message ?? String(e)}`);
		}
	}
	return created;
}

async function createClosedOutline(
	layerId: number,
	pts: Pt[],
	label?: string,
	opts?: { allowLineFallback?: boolean },
): Promise<{ kind: 'polyline' | 'solidRegion' | 'lines'; ids: string[] }> {
	const allowLineFallback = opts?.allowLineFallback !== false;
	let polylineErr: any = null;
	let regionErr: any = null;
	try {
		const id = await createClosedPolyline(layerId, pts, label);
		return { kind: 'polyline', ids: id ? [id] : [] };
	}
	catch (e) {
		polylineErr = e;
		// Some builds refuse polyline creation on 3D Shell layers. Try SolidRegion/Region if available.
		try {
			const region = await createClosedSolidRegion(layerId, pts, label);
			return { kind: 'solidRegion', ids: region ? [region] : [] };
		}
		catch (re) {
			regionErr = re;
		}

		if (!allowLineFallback) {
			const parts: string[] = [];
			if (polylineErr?.message) parts.push(`Polyline: ${polylineErr.message}`);
			else if (polylineErr) parts.push(`Polyline: ${String(polylineErr)}`);
			if (regionErr?.message) parts.push(`Region: ${regionErr.message}`);
			else if (regionErr) parts.push(`Region: ${String(regionErr)}`);
			throw new Error(`${label ? `${label}: ` : ''}3D Shell outline creation failed.\n${parts.join('\n')}`);
		}
		const ids = await createClosedLoopAsLines(layerId, pts, label);
		return { kind: 'lines', ids };
	}
}

async function deleteRegionsOnLayers(layers: number[]) {
	const api: any =
		(eda as any)?.pcb_PrimitiveSolidRegion
		?? (eda as any)?.pcb_PrimitiveRegion
		?? null;
	if (!api?.getAll || !api?.delete) return;
	const all = asArray<any>(await api.getAll());
	const toDelete = all
		.filter((p) => layers.includes(getLayerId(p) ?? -1))
		.map((p) => p?.primitiveId ?? p?.id ?? p?.getState_PrimitiveId?.());
	const ids = toDelete.filter((id: any) => id != null).map((id: any) => String(id));
	if (!ids.length) return;
	try {
		await api.delete(ids);
	}
	catch {
		for (const id of ids) {
			try { await api.delete([id]); } catch {}
		}
	}
}

async function createClosedSolidRegion(layerId: number, pts: Pt[], label?: string): Promise<string | null> {
	// Best-effort: different EasyEDA builds expose different names for "solid region" primitives.
	const api: any =
		(eda as any)?.pcb_PrimitiveSolidRegion
		?? (eda as any)?.pcb_PrimitiveRegion
		?? null;
	if (!api?.create) throw new Error('pcb_PrimitiveSolidRegion/pcb_PrimitiveRegion.create not available');
	if (!pts || pts.length < 3) throw new Error('solid region requires >= 3 points');

	const mathApi: any = (eda as any)?.pcb_MathPolygon;
	if (!mathApi?.createPolygon) throw new Error('pcb_MathPolygon.createPolygon not available');

	const ptsNorm = normalizePoints(pts, true);
	const polygon = await createPolygonRobust(mathApi, ptsNorm);
	if (!polygon) throw new Error('pcb_MathPolygon.createPolygon returned null/undefined');

	const net = '';
	const lock = false;

	const findCreated = async (id: string) => {
		try {
			if (!api.getAll) return null;
			const all = asArray<any>(await api.getAll());
			for (const p of all) {
				const pid =
					(typeof p?.getState_PrimitiveId === 'function' ? p.getState_PrimitiveId() : undefined)
					?? p?.primitiveId ?? p?.id;
				if (pid && String(pid) === id) return p;
			}
		}
		catch {}
		return null;
	};

	const tryMoveToLayer = async (id: string) => {
		try {
			if (!api.modify) return false;
			await api.modify(id, { layerId });
		}
		catch {
			try { await api.modify(id, { layer: layerId }); } catch { return false; }
		}
		const created = await findCreated(id);
		return created ? getLayerId(created) === layerId : false;
	};

	const variants: Array<{ desc: string; args: any[] }> = [
		{ desc: 'net,layer,polygon,lock', args: [net, layerId as any, polygon, lock] },
		{ desc: 'net,polygon,layer,lock', args: [net, polygon, layerId as any, lock] },
		{ desc: 'layer,polygon,lock', args: [layerId as any, polygon, lock] },
		{ desc: 'options', args: [{ net, layerId, layer: layerId, polygon, lock, primitiveLock: lock, name: label }] },
	];

	let lastErr: any = null;
	for (const v of variants) {
		try {
			const res = await api.create.apply(api, v.args);
			const id = getPrimitiveIdFromUnknown(res);
			if (!id) continue;
			const created = await findCreated(id);
			const actualLayer = created ? getLayerId(created) : null;
			if (actualLayer === layerId) return id;
			if (await tryMoveToLayer(id)) return id;
			try { if (api.delete) await api.delete([id]); } catch {}
		}
		catch (e: any) {
			lastErr = e;
		}
	}

	throw new Error(`${label ? `${label}: ` : ''}solid region create failed. Last error: ${lastErr?.message ?? String(lastErr)}`);
}

async function generate3dShell(payload: ShellPayload) {
	const pcbApi: any = eda as any;
	if (!pcbApi?.pcb_PrimitiveComponent?.getAll) {
		throw new Error('PCB API not available. Open a PCB document first.');
	}

	const rawJson = payload?.json;
	if (!rawJson) throw new Error('Missing KLE JSON');

	const pitch = Number(payload.pitch);
	if (!Number.isFinite(pitch) || pitch <= 0) throw new Error('Invalid pitch');

	const outlineClearanceMm = Number.isFinite(Number(payload.outlineClearanceMm)) ? Number(payload.outlineClearanceMm) : 0.2;
	const wallThicknessMm = Number.isFinite(Number(payload.wallThicknessMm)) ? Number(payload.wallThicknessMm) : 2.0;
	const cutoutSizeMm = Number.isFinite(Number(payload.cutoutSizeMm)) ? Number(payload.cutoutSizeMm) : 14.0;
	const applyKeyRotation = payload.applyKeyRotation !== false;
	const createKeyHoles = payload.createKeyHoles !== false;
	const keyHolesTop = payload.keyHolesTop !== false;
	const keyHolesBottom = payload.keyHolesBottom !== false;
	const onlyLegendKeys = Boolean(payload.onlyLegendKeys);
	const clearExistingShellLayers = payload.clearExistingShellLayers !== false;
	const rectUnionMarginMm = Number.isFinite(Number(payload.rectUnionMarginMm)) ? Number(payload.rectUnionMarginMm) : 2.0;
	const outlineSource = String(payload.outlineSource || '') === 'plateHull'
		? 'plateHull'
		: (String(payload.outlineSource || '') === 'rectUnion'
			? 'rectUnion'
			: (String(payload.outlineSource || '') === 'selected' ? 'selected' : 'board'));
	const useSelectedOutline = outlineSource === 'selected' || Boolean(payload.useSelectedOutline);

	const LAYER_BOARD_OUTLINE = 11;
	const LAYER_SHELL_3D_OUTLINE = 53;
	const LAYER_SHELL_3D_TOP = 54;
	const LAYER_SHELL_3D_BOTTOM = 55;

	showToast('KLE Importer Pro: generating 3D Shell (experimental)...', 2);

	if (clearExistingShellLayers) {
		await deletePolylinesOnLayers([LAYER_SHELL_3D_OUTLINE, LAYER_SHELL_3D_TOP, LAYER_SHELL_3D_BOTTOM]);
		await deleteLinesOnLayers([LAYER_SHELL_3D_OUTLINE, LAYER_SHELL_3D_TOP, LAYER_SHELL_3D_BOTTOM]);
		await deleteRegionsOnLayers([LAYER_SHELL_3D_OUTLINE, LAYER_SHELL_3D_TOP, LAYER_SHELL_3D_BOTTOM]);
	}

	// 1) Read BOARD_OUTLINE polyline(s) and pick the largest closed polygon.
	const polyApi: any = (eda as any)?.pcb_PrimitivePolyline;
	let baseOutline: Pt[] | null = null;

	// Use component locations as "samples" to pick the correct outline when multiple loops exist.
	// Prefer SW/D footprints (usually fill the board area).
	let samplePts: Pt[] = [];
	try {
		const compsRaw = await (eda as any).pcb_PrimitiveComponent.getAll();
		const comps = asArray<any>(compsRaw);
		const filtered = comps.filter((c) => {
			const d = String(c?.designator ?? c?.getState_Designator?.() ?? '');
			return /^SW\d+$/i.test(d) || /^D\d+$/i.test(d);
		});
		samplePts = filtered.map(extractComponentPoint).filter((p): p is Pt => Boolean(p));
		// If SW/D didn't provide positions, try any component positions.
		if (samplePts.length < 3) {
			samplePts = comps.map(extractComponentPoint).filter((p): p is Pt => Boolean(p));
		}
	}
	catch {}

	// 1--1) Alternative outline source: use the same logic as "switch plate export"
	// (convex hull of switch cutouts + margin offsets).
	if (outlineSource === 'plateHull') {
		const kle = parseKleText(rawJson);
		const milPerU = (pitch / 25.4) * 1000;
		const layout = parseKLE(kle, 0, 0);
		const sizeMil = mmToMil(cutoutSizeMm);
		const vertices: Pt[] = [];
		for (const k of layout) {
			if (onlyLegendKeys) {
				const hint = k?.swHint;
				if (!(typeof hint === 'number' && Number.isFinite(hint))) continue;
			}
			const cx = Number(k?.swX) * milPerU;
			const cy = -Number(k?.swY) * milPerU;
			const rot = -Number(k?.rot || 0);
			if (!Number.isFinite(cx) || !Number.isFinite(cy)) continue;
			const rect = rectPointsMil(cx, cy, sizeMil, rot, applyKeyRotation);
			for (const p of rect) vertices.push(p);
		}
		const hull = convexHull(vertices);
		if (hull.length < 3) {
			throw new Error('plateHull outline failed: not enough points from KLE keys.');
		}
		baseOutline = hull;
	}

	// 1--2) Alternative outline source: boolean union of per-key rectangles (more detailed than convex hull).
	if (outlineSource === 'rectUnion') {
		const kle = parseKleText(rawJson);
		const milPerU = (pitch / 25.4) * 1000;
		const layout = parseKLE(kle, 0, 0);
		const sizeMil = mmToMil(cutoutSizeMm + 2 * rectUnionMarginMm);
		const rects: Pt[][] = [];
		for (const k of layout) {
			if (onlyLegendKeys) {
				const hint = k?.swHint;
				if (!(typeof hint === 'number' && Number.isFinite(hint))) continue;
			}
			const cx = Number(k?.swX) * milPerU;
			const cy = -Number(k?.swY) * milPerU;
			const rot = -Number(k?.rot || 0);
			if (!Number.isFinite(cx) || !Number.isFinite(cy)) continue;
			rects.push(rectPointsMil(cx, cy, sizeMil, rot, applyKeyRotation));
		}
		const union = unionPolygonsClipper(rects);
		baseOutline = pickBestOutline(union, samplePts);
		if (!baseOutline) {
			throw new Error('rectUnion outline failed: union produced no polygons.');
		}
	}

	// 1-0) Optional: use selected primitives as the outline source (avoids picking the wrong shapes).
	if (!baseOutline && useSelectedOutline && (eda as any)?.pcb_SelectControl?.getAllSelectedPrimitives_PrimitiveId) {
		const selectedIds = asArray<any>(await (eda as any).pcb_SelectControl.getAllSelectedPrimitives_PrimitiveId());
		const selectedSet = new Set<string>(selectedIds.filter((id: any) => typeof id === 'string').map((id: string) => String(id)));

		const isSelected = (p: any) => {
			const pid =
				(typeof p?.getState_PrimitiveId === 'function' ? p.getState_PrimitiveId() : undefined)
				?? p?.primitiveId ?? p?.id;
			return pid && selectedSet.has(String(pid));
		};

		if (selectedSet.size) {
			// Try selected polyline first
			if (polyApi?.getAll) {
				const allPolys = asArray<any>(await polyApi.getAll());
				const selPolys = allPolys.filter(isSelected);
				const candidates = selPolys
					.map((p) => extractPolylinePoints(p))
					.filter((pts) => pts && pts.length >= 3) as Pt[][];
				baseOutline = pickBestOutline(candidates, samplePts);
			}

			// Else try selected lines/arcs
			if (!baseOutline) {
				const lineApi: any = (eda as any)?.pcb_PrimitiveLine;
				const arcApi: any = (eda as any)?.pcb_PrimitiveArc;
				const lines = lineApi?.getAll ? asArray<any>(await lineApi.getAll()) : [];
				const arcs = arcApi?.getAll ? asArray<any>(await arcApi.getAll()) : [];

				const items: OutlineItem[] = [];
				for (const l of lines) {
					if (!isSelected(l)) continue;
					const it = extractLineItem(l);
					if (it) items.push(it);
				}
				for (const a of arcs) {
					if (!isSelected(a)) continue;
					const it = extractArcItem(a);
					if (it) items.push(it);
				}

				const cycles = extractCyclesFromItems(items, 0.01);
				if (cycles.length) {
					baseOutline = pickBestOutline(cycles, samplePts);
				}
				else {
					const path = buildClosedPathFromItems(items, 0.01);
					if (path && path.length >= 3) baseOutline = path;
				}
			}

			if (!baseOutline) {
				throw new Error('useSelectedOutline is enabled, but selected primitives do not form a closed outline.');
			}
		}
	}

	// 1-a) Try BOARD_OUTLINE polylines (best case)
	if (!baseOutline && polyApi?.getAll) {
		const allPolys = asArray<any>(await polyApi.getAll());
		const outlineCandidates = allPolys
			.filter((p) => getLayerId(p) === LAYER_BOARD_OUTLINE)
			.map((p) => extractPolylinePoints(p))
			.filter((pts) => pts && pts.length >= 3) as Pt[][];

		baseOutline = pickBestOutline(outlineCandidates, samplePts);
	}

	// 1-b) Fallback: BOARD_OUTLINE made of lines/arcs
	if (!baseOutline) {
		const lineApi: any = (eda as any)?.pcb_PrimitiveLine;
		const arcApi: any = (eda as any)?.pcb_PrimitiveArc;
		const lines = lineApi?.getAll ? asArray<any>(await lineApi.getAll()) : [];
		const arcs = arcApi?.getAll ? asArray<any>(await arcApi.getAll()) : [];

		const items: OutlineItem[] = [];
		for (const l of lines) {
			if (getLayerId(l) !== LAYER_BOARD_OUTLINE) continue;
			const it = extractLineItem(l);
			if (it) items.push(it);
		}
		for (const a of arcs) {
			if (getLayerId(a) !== LAYER_BOARD_OUTLINE) continue;
			const it = extractArcItem(a);
			if (it) items.push(it);
		}

		const cycles = extractCyclesFromItems(items, 0.01);
		if (cycles.length) {
			baseOutline = pickBestOutline(cycles, samplePts);
		}
		else {
			const path = buildClosedPathFromItems(items, 0.01);
			if (path && path.length >= 3) baseOutline = path;
		}
	}

	if (!baseOutline) {
		throw new Error(
			'BOARD_OUTLINE (layer 11) not found.\n' +
			'This feature needs a board outline on layer 11 (polyline or a closed loop of lines/arcs).'
		);
	}

	// Heuristic warning: if the chosen outline is wildly larger than component spread,
	// it's likely we picked an unrelated loop. Recommend "use selected outline".
	try {
		const sb = bboxOfPoints(samplePts);
		if (sb) {
			const sArea = bboxAreaFromBounds(sb);
			const oArea = bboxArea(baseOutline);
			if (sArea > 1 && oArea > sArea * 80) {
				showToast('3D Shell: outline looks too large. Try enabling "use selected outline".', 4);
			}
		}
	}
	catch {}

	// 2) Offset for shell wall region.
	const clearanceMil = mmToMil(outlineClearanceMm);
	const outerMil = mmToMil(outlineClearanceMm + wallThicknessMm);

	const inner = offsetPolygonClipper(baseOutline, clearanceMil);
	const outer = offsetPolygonClipper(baseOutline, outerMil);
	if (!inner || !outer) {
		throw new Error('Failed to offset BOARD_OUTLINE. (Unsupported outline shape?)');
	}

	// 3) Create outlines on SHELL_3D_OUTLINE layer (two paths: outer+inner).
	const outerCreated = await createClosedOutline(LAYER_SHELL_3D_OUTLINE, outer, 'shell_outer_outline', { allowLineFallback: false });
	const innerCreated = await createClosedOutline(LAYER_SHELL_3D_OUTLINE, inner, 'shell_inner_outline', { allowLineFallback: false });

	// 4) Create key holes (rectangles) on SHELL_3D_TOP/BOTTOM layers.
	let holesCreated = 0;
	if (createKeyHoles && (keyHolesTop || keyHolesBottom)) {
		const kle = parseKleText(rawJson);
		const milPerU = (pitch / 25.4) * 1000;
		const layout = parseKLE(kle, 0, 0);
		const sizeMil = mmToMil(cutoutSizeMm);

		for (const k of layout) {
			if (onlyLegendKeys) {
				const hint = k?.swHint;
				if (!(typeof hint === 'number' && Number.isFinite(hint))) continue;
			}
			const cx = Number(k?.swX) * milPerU;
			const cy = -Number(k?.swY) * milPerU;
			const rot = -Number(k?.rot || 0);
			if (!Number.isFinite(cx) || !Number.isFinite(cy)) continue;
			const rect = rectPointsMil(cx, cy, sizeMil, rot, applyKeyRotation);
			if (keyHolesTop) {
				await createClosedOutline(LAYER_SHELL_3D_TOP, rect, 'key_hole_top', { allowLineFallback: false });
				holesCreated++;
			}
			if (keyHolesBottom) {
				await createClosedOutline(LAYER_SHELL_3D_BOTTOM, rect, 'key_hole_bottom', { allowLineFallback: false });
				holesCreated++;
			}
		}
	}

	showInfo(
		`3D Shell primitives created.\n` +
		`- Layer ${LAYER_SHELL_3D_OUTLINE}: outer+inner outlines (outer=${outerCreated.kind}, inner=${innerCreated.kind})\n` +
		`- Layer ${LAYER_SHELL_3D_TOP}/${LAYER_SHELL_3D_BOTTOM}: key holes = ${holesCreated}\n\n` +
		`Next: open EasyEDA's 3D Shell tool to preview/export.\n` +
		`(Note: behavior depends on EasyEDA's 3D Shell implementation; this is experimental.)`,
		'KLE Importer Pro'
	);

	return {
		ok: true,
		created: {
			outline: 2,
			keyHoles: holesCreated,
		},
		params: {
			outlineClearanceMm,
			wallThicknessMm,
			cutoutSizeMm,
		},
	};
}
*/

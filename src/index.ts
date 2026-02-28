import * as extensionConfig from '../extension.json';

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
		await sysIFrame.openIFrame('/iframe/index.html', 380, 470, 'kle-importer-pro', {
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

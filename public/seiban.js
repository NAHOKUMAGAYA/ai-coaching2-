// 製番転記.ps1 と同じ処理をブラウザの中だけで行う（ファイルはどこにも送らない）
// 注残リストの「納入指示番号」から製番台帳を引いて「FD製番」「部品納期」列に書き込み、
// その製番の発注明細(発注仕入状況一覧表)から一番遅い回答納期を「納期回答」列に、回答備考を「コメント」列に書き込む
(function (global) {
    'use strict';

    const SKIP_SHEETS = ['記入方法', 'リスト', '休日'];
    // 納期回答・回答備考を見るときに無視する発注先番号（この発注先の明細は回答納期が入っていても使わない）
    const IGNORE_SUPPLIERS = ['FSH', 'FZH', 'FZH2', 'NK', 'SFD', 'SFH', 'SFZ', 'SM', 'SMF', 'SPH', 'SPM', 'SPP', 'STK'];
    // FD製番に「対象外」と入れる条件（どれか1つに当てはまる行）:
    //   ユーザ欄1が J 以外 / 納入指示区分が JIT（手配JIT は含まない）/ 品番が対象外品番リストのこのシートにある
    const EXCLUDE_TEXT = '対象外';
    const EXCLUDE_SHEET = '納期回答対象外';
    // 上の条件に当てはまっても対象外にしない品番
    const KEEP_PARTS = ['EK0393'];

    // セルの色分け
    //   製番が2つ以上 : 黄色塗り・赤太字（ほかより優先）
    //   納期回答の新規: 水色塗り・白字 / 変更: 赤塗り・白字 / 前回と同じ: 塗りなし・黒字
    const MULTI_FILL = 'FFFFFF00';
    const MULTI_FONT = 'FFFF0000';
    const CHANGED_FILL = 'FFFF0000';
    const ADDED_FILL = 'FF00B0F0';
    const MARK_FONT = 'FFFFFFFF';
    const HEAD_FILL = 'FFD9D9D9';

    // 読むだけのファイルは、使わない部分を読み飛ばして速くする
    const READ_ONLY_OPTIONS = { ignoreNodes: ['dataValidations', 'conditionalFormatting', 'hyperlinks', 'drawing', 'picture', 'tableParts', 'sheetProtection', 'rowBreaks', 'headerFooter', 'printOptions', 'pageMargins', 'pageSetup'] };

    const DAY_MS = 86400000;

    function str(v) {
        return v === null || v === undefined ? '' : String(v);
    }

    function normalizeKey(s) {
        s = str(s);
        if (!s) return '';
        return s.normalize('NFKC').trim().toUpperCase();
    }

    // Excel の Value2 と同じ形にする（日付は数値のシリアル値、結合セルの左上以外は空）
    function cellVal(cell) {
        if (cell.isMerged && cell.master !== cell) return null;
        return plain(cell.value);
    }

    function plain(v) {
        if (v === null || v === undefined) return null;
        if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
        if (v instanceof Date) return Math.round((v.getTime() / DAY_MS + 25569) * 86400) / 86400;
        if (v.richText) return v.richText.map(t => t.text).join('');
        if (v.formula !== undefined || v.sharedFormula !== undefined) return plain(v.result);
        if (v.hyperlink !== undefined) return plain(v.text);
        return null;   // エラー値など
    }

    function isNumber(v) {
        return typeof v === 'number';
    }

    function pad(n, w) {
        return String(n).padStart(w || 2, '0');
    }

    // 台帳の日付(シリアル値)を注残リストの他の日付列と同じ yyyy-MM-dd の文字にする
    function formatDate(v) {
        if (isNumber(v)) {
            const d = new Date(Math.round((v - 25569) * DAY_MS));
            return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
        }
        return str(v).trim();
    }

    function formatNow(d) {
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }

    function stamp(d) {
        return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
    }

    function formatDays(days) {
        if (days === '' || days === null || days === undefined) return '';
        if (days > 0) return `+${days}日遅れ`;
        if (days < 0) return `${-days}日早まり`;
        return '';
    }

    function parseDay(s) {
        const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || '');
        if (!m) return null;
        const t = Date.UTC(+m[1], +m[2] - 1, +m[3]);
        const d = new Date(t);
        return d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3] ? t : null;
    }

    // 1～maxRow 行目・1～maxCol 列目から、指定パターンの見出しがすべて揃う行を探す
    function findHeader(ws, maxRow, maxCol, patterns) {
        const names = Object.keys(patterns);
        for (let r = 1; r <= maxRow; r++) {
            const found = {};
            let n = 0;
            for (let c = 1; c <= maxCol; c++) {
                const t = str(cellVal(ws.getCell(r, c))).replace(/\s/g, '');
                for (const name of names) {
                    if (!(name in found) && patterns[name].test(t)) { found[name] = c; n++; }
                }
            }
            if (n === names.length) { found.row = r; return found; }
        }
        return null;
    }

    // 列の一番下の値がある行（Excel の End(xlUp) と同じ。何も無ければ 1）
    function lastRow(ws, col) {
        for (let r = ws.rowCount; r > 1; r--) {
            const cell = ws.findCell(r, col);
            if (!cell || (cell.isMerged && cell.master !== cell)) continue;
            if (cell.value !== null && cell.value !== undefined && cell.value !== '') return r;
        }
        return 1;
    }

    function readColumn(ws, col, last) {
        const out = new Array(last + 1).fill(null);
        for (let r = 1; r <= last; r++) {
            const cell = ws.findCell(r, col);
            if (cell) out[r] = cellVal(cell);
        }
        return out;
    }

    function addValue(map, key, value) {
        if (!value) return;
        if (!map.has(key)) map.set(key, []);
        const list = map.get(key);
        if (!list.includes(value)) list.push(value);
    }

    function setStyle(cell, fill, fontColor, bold) {
        cell.fill = fill ? { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } } : { type: 'pattern', pattern: 'none' };
        const font = Object.assign({}, cell.font);
        delete font.color;
        if (fontColor) font.color = { argb: fontColor };
        font.bold = bold;
        cell.font = font;
    }

    function setMultiStyle(cell) { setStyle(cell, MULTI_FILL, MULTI_FONT, true); }
    function setPlainStyle(cell) { setStyle(cell, null, null, false); }

    // 製番や日付が数値に変換されないよう文字列書式にして書き込む
    function setText(cell, value) {
        cell.numFmt = '@';
        cell.value = value === '' ? null : value;
    }

    function clearNote(cell) {
        cell._comment = undefined;
    }

    async function loadBook(data, options) {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(data, options);
        return wb;
    }

    // ---- 製番台帳: 納入指示番号 → 製番・部品納期 ----
    async function readLedgers(ledgers, log) {
        const map = new Map();
        const dateMap = new Map();
        for (const ledger of ledgers) {
            log(`台帳読込: ${ledger.name}`);
            let wb;
            try {
                wb = await loadBook(await ledger.read(), READ_ONLY_OPTIONS);
            } catch (e) {
                log(`  ※読み込めません。スキップします（${e.message}）`);
                continue;
            }
            for (const ws of wb.worksheets) {
                if (SKIP_SHEETS.includes(ws.name)) continue;
                const h = findHeader(ws, 10, 36, { seiban: /^製番$/i, order: /客先O\/N/i });
                if (!h) continue;
                const last = lastRow(ws, h.seiban);
                if (last <= h.row) continue;
                const sb = readColumn(ws, h.seiban, last);
                const on = readColumn(ws, h.order, last);
                let dateCol = 0;
                for (let c = 1; c <= 36; c++) {
                    if (str(cellVal(ws.getCell(h.row, c))).replace(/\s/g, '') === '部品納期') { dateCol = c; break; }
                }
                const dt = dateCol ? readColumn(ws, dateCol, last) : null;
                let n = 0;
                for (let r = h.row + 1; r <= last; r++) {
                    const seiban = str(sb[r]).trim();
                    const orders = str(on[r]);
                    if (!seiban || !orders) continue;
                    const date = dateCol ? formatDate(dt[r]) : '';
                    for (let k of orders.split(/[,、，;；\s]+/)) {
                        k = normalizeKey(k);
                        if (!k) continue;
                        addValue(map, k, seiban);
                        addValue(dateMap, k, date);
                        n++;
                    }
                }
                log(`  シート[${ws.name}] 納入指示番号 ${n} 件${dateCol ? '' : '（部品納期の列なし）'}`);
            }
        }
        return { map, dateMap };
    }

    // 発注仕入状況一覧表の明細のうち「発注明細」シートに出す列（見出し → 発注明細シートでの見出し）
    const DETAIL_COLUMNS = [
        ['発注番号', '発注番号'], ['発注先番号', '発注先番号'], ['発注先名称', '発注先名称'], ['品名規格', '品名規格'], ['型式材質', '型式材質'],
        ['発注数計', '発注数'], ['仕入数計', '仕入数'], ['仕入残', '仕入残'], ['仕入状況', '仕入状況'],
        ['発注日', '発注日'], ['部品納期', '部品納期'], ['回答納期', '回答納期']
    ];

    // 発注仕入状況一覧表から 製造番号 → 一番遅い回答納期・未回答の明細の有無 を集計する。
    // detail には「発注明細」シート用に、製造番号ごとの明細の行番号と各列の値を入れる
    async function readAnswers(order, log, detail) {
        log(`発注仕入状況一覧表読込: ${order.name}`);
        const answers = new Map();
        const wb = await loadBook(await order.read(), READ_ONLY_OPTIONS);
        const ws = wb.worksheets[0];
        const h = findHeader(ws, 5, 78, { supplier: /^発注先番号$/i, seiban: /^製造番号$/i, answer: /^回答納期$/i, remark: /^回答備考$/i, remain: /^仕入残$/i });
        if (!h) throw new Error('発注仕入状況一覧表に「発注先番号」「製造番号」「回答納期」「回答備考」「仕入残」の見出しが見つかりません。');
        const last = lastRow(ws, h.seiban);
        const sp = readColumn(ws, h.supplier, last);
        const sb = readColumn(ws, h.seiban, last);
        const an = readColumn(ws, h.answer, last);
        const bk = readColumn(ws, h.remark, last);
        const rm = readColumn(ws, h.remain, last);

        detail.cols = new Map();
        for (const [name] of DETAIL_COLUMNS) {
            for (let c = 1; c <= 80; c++) {
                if (str(cellVal(ws.getCell(h.row, c))).replace(/\s/g, '') === name) {
                    detail.cols.set(name, readColumn(ws, c, last));
                    break;
                }
            }
        }
        detail.rows = new Map();
        detail.remarks = new Array(last + 1).fill('');

        let prevRemark = '';
        let ignored = 0;
        for (let r = h.row + 1; r <= last; r++) {
            // 「〃」はすぐ上の行(同じ発注先)の回答備考と同じという意味なので置き換える
            let remark = str(bk[r]).trim();
            if (/^[〃″"]+$/.test(remark) && r > h.row + 1 && str(sp[r]) === str(sp[r - 1])) remark = prevRemark;
            prevRemark = remark;
            detail.remarks[r] = remark;

            const k = normalizeKey(sb[r]);
            if (!k) continue;
            if (!detail.rows.has(k)) detail.rows.set(k, []);
            detail.rows.get(k).push(r);
            if (IGNORE_SUPPLIERS.includes(normalizeKey(sp[r]))) { ignored++; continue; }
            if (!answers.has(k)) {
                answers.set(k, {
                    max: null, pending: false,
                    maxRemarks: [],      // 一番遅い回答納期の明細の回答備考
                    pendingRemarks: []   // 回答納期が空で入荷が残る明細の回答備考
                });
            }
            const e = answers.get(k);
            const a = an[r];
            if (isNumber(a)) {
                if (e.max === null || a > e.max) { e.max = a; e.maxRemarks = []; }
                if (a === e.max && remark && !e.maxRemarks.includes(remark)) e.maxRemarks.push(remark);
            } else if (isNumber(rm[r]) && rm[r] > 0) {
                // 入荷が残っているのに回答納期が空の明細
                e.pending = true;
                if (remark && !e.pendingRemarks.includes(remark)) e.pendingRemarks.push(remark);
            }
        }
        log(`  発注明細 ${last - h.row} 行（うち無視する発注先 ${ignored} 行） / 製造番号 ${answers.size} 件`);
        return answers;
    }

    // 対象外品番リストの「納期回答対象外」シートから品番を読む
    async function readExcludeItems(source, log) {
        log(`対象外品番リスト読込: ${source.name}`);
        const items = new Set();
        const wb = await loadBook(await source.read(), READ_ONLY_OPTIONS);
        const ws = wb.getWorksheet(EXCLUDE_SHEET);
        if (!ws) throw new Error(`対象外品番リストに「${EXCLUDE_SHEET}」シートが見つかりません。`);
        const h = findHeader(ws, 5, 26, { item: /^品番$/i });
        if (!h) throw new Error(`対象外品番リストの「${EXCLUDE_SHEET}」シートに「品番」の見出しが見つかりません。`);
        const last = lastRow(ws, h.item);
        const v = readColumn(ws, h.item, last);
        for (let r = h.row + 1; r <= last; r++) {
            const k = normalizeKey(v[r]);
            if (k) items.add(k);
        }
        log(`  シート[${EXCLUDE_SHEET}] 品番 ${items.size} 件`);
        return items;
    }

    // FD製番に「対象外」と入れる行を決める。ユーザ欄1が J 以外・納入指示区分が JIT・品番が対象外品番リストにある、のどれかに当てはまる行。
    // 見出しが無い条件は使わない
    function findExcluded(ws, first, count, items, log) {
        const col = (pattern, label) => {
            const h = findHeader(ws, 5, 78, { c: pattern });
            if (!h) { log(`※注残リストに「${label}」の見出しが無いため、この条件では対象外にしません`); return null; }
            return h.c;
        };
        const userCol = col(/^ユーザ欄[1１]$/, 'ユーザ欄1');
        const kindCol = col(/^納入指示区分$/, '納入指示区分');
        const partCol = col(/^品番$/, '品番');
        const get = (r, c) => normalizeKey(cellVal(ws.getCell(r, c)));

        const stat = { flags: new Array(count).fill(false), rows: 0, user: 0, jit: 0, item: 0 };
        for (let i = 0; i < count; i++) {
            const row = first + i;
            let hit = false;
            const part = partCol ? get(row, partCol) : '';
            if (part && KEEP_PARTS.includes(part)) continue;
            if (userCol && get(row, userCol) !== 'J') { stat.user++; hit = true; }
            // 「1:ＪＩＴ」のように番号が付いているので外して比べる
            if (kindCol && get(row, kindCol).replace(/^\d+:/, '') === 'JIT') { stat.jit++; hit = true; }
            if (items && part && items.has(part)) { stat.item++; hit = true; }
            if (hit) { stat.flags[i] = true; stat.rows++; }
        }
        return stat;
    }

    // 注残リストに書き込む列（見出し, 探すときのパターン）。無い列は納入指示番号の右にこの順で挿入する
    const FILL_COLUMNS = [['FD製番', /^FD製番$/i], ['部品納期', /^部品納期$/i], ['納期回答', /^納期回答$/i], ['コメント', /^(コメント|回答備考)$/i]];

    function addMissingColumns(ws, log) {
        const h = findHeader(ws, 5, 78, { order: /^納入指示番号$/i });
        if (!h) return;   // 見出しが無いエラーはこのあとで出す
        let col = h.order;
        const added = [];
        for (const [name, pattern] of FILL_COLUMNS) {
            const found = findHeader(ws, 5, 78, { c: pattern });
            if (found) { col = found.c; continue; }
            col++;
            ws.spliceColumns(col, 0, []);
            // 見出しの書式は左隣(納入指示番号など)と同じにする
            const head = ws.getCell(h.row, col);
            head.style = JSON.parse(JSON.stringify(ws.getCell(h.row, col - 1).style || {}));
            head.value = name;
            ws.getColumn(col).width = 12;
            widenAutoFilter(ws, col);
            added.push(name);
        }
        if (added.length) log(`注残リストに列を挿入しました: ${added.join('・')}`);
    }

    // 列を挿入したぶん、オートフィルタの範囲を右へ広げる（ExcelJS は自動では直さない）
    function widenAutoFilter(ws, col) {
        const m = typeof ws.autoFilter === 'string' && /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/.exec(ws.autoFilter);
        if (!m) return;
        const end = ws.getColumn(m[3]).number;
        if (col <= end) ws.autoFilter = `${m[1]}${m[2]}:${ws.getColumn(end + 1).letter}${m[4]}`;
    }

    // ---- CSV ----
    function parseCsv(text) {
        text = text.replace(/^﻿/, '');
        const rows = [];
        let row = [], field = '', quoted = false;
        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            if (quoted) {
                if (ch === '"') {
                    if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false;
                } else field += ch;
            } else if (ch === '"') quoted = true;
            else if (ch === ',') { row.push(field); field = ''; }
            else if (ch === '\n' || ch === '\r') {
                if (ch === '\r' && text[i + 1] === '\n') i++;
                row.push(field); rows.push(row); row = []; field = '';
            } else field += ch;
        }
        if (field !== '' || row.length) { row.push(field); rows.push(row); }
        const head = rows.shift() || [];
        return rows.filter(r => r.length > 1 || r[0] !== '').map(r => {
            const o = {};
            head.forEach((name, i) => { o[name] = r[i]; });
            return o;
        });
    }

    // PowerShell の Export-Csv と同じ形（UTF-8 BOM 付き・全部ダブルクォート）
    function toCsv(heads, rows) {
        const q = v => '"' + str(v).replace(/"/g, '""') + '"';
        const lines = [heads.map(q).join(',')];
        for (const r of rows) lines.push(heads.map(h => q(r[h])).join(','));
        return '﻿' + lines.join('\r\n') + '\r\n';
    }

    // 前回までに入れた納期回答の記録（納入指示番号ごと）
    function readHistory(text) {
        const history = new Map();
        if (!text) return history;
        for (const row of parseCsv(text)) {
            history.set(str(row['納入指示番号']), { value: str(row['納期回答']), seiban: str(row['FD製番']), time: str(row['記録日時']), remark: str(row['回答備考']) });
        }
        return history;
    }

    function historyCsv(history) {
        const keys = [...history.keys()].sort((a, b) => a.localeCompare(b, 'ja'));
        return toCsv(['納入指示番号', 'FD製番', '納期回答', '記録日時', '回答備考'], keys.map(k => {
            const h = history.get(k);
            return { '納入指示番号': k, 'FD製番': h.seiban, '納期回答': h.value, '記録日時': h.time, '回答備考': h.remark };
        }));
    }

    // 「対象外」は製番ではないので外す
    function splitSeiban(v) {
        return str(v).split('/').map(normalizeKey).filter(s => s && s !== EXCLUDE_TEXT);
    }

    // 注残リストの1列を、納入指示番号をキーに対応表から埋める。候補が複数ある行は「/」でつないで黄色塗り・赤太字にする。
    // 対象外の行(excluded)は台帳を引かず、excludedText があればそれを入れる（無ければ空のまま）
    function fillColumn(ws, first, col, keys, lookup, overwrite, excluded, excludedText) {
        const stat = { filled: 0, multi: 0, kept: 0, excluded: 0 };
        keys.forEach((k, i) => {
            const cell = ws.getCell(first + i, col);
            const existing = str(cellVal(cell));
            if (!overwrite && existing !== '') {
                // 残した行も、前の色分けやダウンロード時の赤太字のままにならないよう色分けし直す
                stat.kept++;
                if (existing.includes('/')) setMultiStyle(cell); else setPlainStyle(cell);
                return;
            }
            if (excluded[i]) {
                if (excludedText) { setText(cell, excludedText); stat.excluded++; setPlainStyle(cell); }
                return;
            }
            if (!k || !lookup.has(k)) return;
            const values = lookup.get(k);
            setText(cell, values.join('/'));
            stat.filled++;
            if (values.length > 1) { stat.multi++; setMultiStyle(cell); } else setPlainStyle(cell);
        });
        return stat;
    }

    // FD製番ごとに一番遅い回答納期を「納期回答」列に毎回入れ直す。FD製番が「/」で複数ならその中で一番遅い日。
    // 前回の記録(無ければセルにあった値)と変わった行は赤塗り、空欄から入った行は水色塗りにして、内容をメモに残す
    function fillAnswer(ws, first, count, seibanCol, col, answers, keys, history, now) {
        const changes = [];
        const pendingRows = [];
        const multiRows = [];
        const stat = { filled: 0, pending: 0, multi: 0, noAnswer: 0, notFound: 0, cleared: 0, changes };

        for (let i = 0; i < count; i++) {
            const row = first + i;
            const cell = ws.getCell(row, col);
            const existing = str(cellVal(cell));
            // 前回の色・文字色・メモを消してから付け直す
            setPlainStyle(cell);
            clearNote(cell);
            const seibanText = str(cellVal(ws.getCell(row, seibanCol)));
            const list = splitSeiban(seibanText);
            if (list.length === 0) continue;
            let max = null, pending = false, found = false;
            for (const s of list) {
                if (!answers.has(s)) continue;
                found = true;
                const e = answers.get(s);
                if (e.pending) pending = true;
                if (e.max !== null && (max === null || e.max > max)) max = e.max;
            }
            const k = keys[i];
            if (!found || max === null) {
                if (!found) stat.notFound++; else stat.noAnswer++;
                // 回答が無くなった行は、前回このツールが入れた値のままなら消す（無視する発注先の回答などが残らないように）。
                // 自分で入力した値(履歴と違う値)は残す
                if (k && existing !== '' && history.has(k) && history.get(k).value === existing) {
                    setText(cell, '');
                    history.get(k).value = '';
                    history.get(k).time = now;
                    stat.cleared++;
                }
                continue;
            }
            const value = formatDate(max);
            setText(cell, value);
            stat.filled++;
            if (pending) { stat.pending++; pendingRows.push(row); }
            if (list.length > 1) { stat.multi++; multiRows.push(row); }

            if (!k) continue;
            let prev = existing, prevTime = '';
            if (history.has(k)) { prev = history.get(k).value; prevTime = history.get(k).time; }
            let kind = '';
            if (!prev) kind = '新規';
            else if (prev !== value) kind = '変更';
            if (kind) {
                const p = parseDay(prev), n = parseDay(value);
                const days = p !== null && n !== null ? Math.round((n - p) / DAY_MS) : '';
                changes.push({ kind, row, key: k, seiban: seibanText, prev, value, days, prevTime, pending, multi: list.length > 1 });
            }
            if (!history.has(k)) history.set(k, { remark: '' });
            const h = history.get(k);
            h.value = value;
            h.seiban = seibanText;
            h.time = now;
        }

        // 新規は水色塗り・白字、変更は赤塗り・白字。製番が複数の行は黄色塗り・赤太字を優先し、新規・変更はメモで分かるようにする
        const notes = new Map();
        for (const c of changes) {
            if (!c.multi) setStyle(ws.getCell(c.row, col), c.kind === '変更' ? CHANGED_FILL : ADDED_FILL, MARK_FONT, false);
            if (c.kind === '変更') {
                const when = c.prevTime ? `（前回記録 ${c.prevTime}）` : '';
                notes.set(c.row, `納期回答が変わりました\n前回 ${c.prev} → 今回 ${c.value} ${formatDays(c.days)}${when}`);
            } else {
                notes.set(c.row, `納期回答が新しく入りました\n今回 ${c.value}`);
            }
        }
        const addNote = (r, text) => notes.set(r, [notes.get(r), text].filter(t => t).join('\n'));
        for (const r of multiRows) { setMultiStyle(ws.getCell(r, col)); addNote(r, '※FD製番が複数あります'); }
        for (const r of pendingRows) addNote(r, '※回答納期が未入力の発注明細があります');
        for (const [r, text] of notes) ws.getCell(r, col).note = text;
        return stat;
    }

    // FD製番ごとの回答備考を「コメント」列に毎回入れ直す。
    // 一番遅い回答納期の明細の備考に、回答納期が空で入荷が残る明細の備考を続ける。
    // 備考が無くなった行は、前回このツールが入れた値のままなら消し、自分で入力した値は残す
    function fillRemark(ws, first, count, seibanCol, col, answers, keys, history, now) {
        const stat = { filled: 0, cleared: 0 };
        for (let i = 0; i < count; i++) {
            const row = first + i;
            const k = keys[i];
            const cell = ws.getCell(row, col);
            // 列を挿入したときに左隣(納期回答)からコピーされた色などが残らないよう、書式を戻す。
            // 「7/15」などが日付に変換されないよう文字列書式にする
            setPlainStyle(cell);
            cell.numFmt = '@';
            const seibanText = str(cellVal(ws.getCell(row, seibanCol)));
            const list = splitSeiban(seibanText).filter(s => answers.has(s));
            let max = null;
            for (const s of list) { const m = answers.get(s).max; if (m !== null && (max === null || m > max)) max = m; }
            const remarks = [];
            const add = t => { if (!remarks.includes(t)) remarks.push(t); };
            for (const s of list) { if (max !== null && answers.get(s).max === max) answers.get(s).maxRemarks.forEach(add); }
            for (const s of list) answers.get(s).pendingRemarks.forEach(add);

            if (remarks.length === 0) {
                const existing = str(cellVal(cell));
                if (k && existing !== '' && history.has(k) && history.get(k).remark === existing) {
                    setText(cell, '');
                    history.get(k).remark = '';
                    stat.cleared++;
                }
                continue;
            }
            const text = remarks.join(' / ');
            setText(cell, text);
            stat.filled++;
            if (k) {
                if (!history.has(k)) history.set(k, { value: '', seiban: seibanText, time: now });
                history.get(k).remark = text;
            }
        }
        return stat;
    }

    // 文字幅（全角は2）で列幅を決める
    function textWidth(s) {
        let w = 0;
        for (const ch of str(s)) w += /[\u0000-ÿ｡-ﾟ]/.test(ch) ? 1 : 2;
        return w;
    }

    // 注残リストの行ごとに、FD製番の発注明細を1行ずつ「発注明細」シートに並べる（毎回作り直す）。
    // 納期回答の元になった明細には★、無視する発注先の明細には「対象外」を付ける
    function writeDetailSheet(wb, ws, first, count, seibanCol, orderCol, answerCol, answers, detail) {
        const sheetName = '発注明細';
        const info = findHeader(ws, 5, 78, { office: /^事業所名$/i, item: /^品名$/i });
        const get = (r, c) => cellVal(ws.getCell(r, c));

        // 出す行（注残の行, 発注明細の行, その注残の行の一番遅い回答納期）を先に集める
        const pairs = [];
        for (let i = 0; i < count; i++) {
            const list = splitSeiban(get(first + i, seibanCol));
            let max = null;
            for (const s of list) {
                if (answers.has(s)) { const m = answers.get(s).max; if (m !== null && (max === null || m > max)) max = m; }
            }
            for (const s of list) {
                if (!detail.rows.has(s)) continue;
                for (const r of detail.rows.get(s)) pairs.push([first + i, r, max]);
            }
        }

        const heads = ['注残の行', '事業所名', '納入指示番号', '品名', 'FD製番', '納期回答', ...DETAIL_COLUMNS.map(d => d[1]), '回答備考', '区分'];
        const numberHeads = ['注残の行', '発注数', '仕入数', '仕入残'];
        const dateHeads = ['発注日', '部品納期', '回答納期'];
        const rows = [heads];
        for (const [row, r, max] of pairs) {
            const out = [
                row,
                info ? str(get(row, info.office)) : '',
                str(get(row, orderCol)),
                info ? str(get(row, info.item)) : '',
                str(get(row, seibanCol)),
                str(get(row, answerCol))
            ];
            for (const [name, head] of DETAIL_COLUMNS) {
                const v = detail.cols.has(name) ? detail.cols.get(name)[r] : null;
                out.push(dateHeads.includes(head) ? formatDate(v) : numberHeads.includes(head) ? v : str(v).trim());
            }
            out.push(detail.remarks[r]);
            const supplier = detail.cols.has('発注先番号') ? normalizeKey(detail.cols.get('発注先番号')[r]) : '';
            const answer = detail.cols.has('回答納期') ? detail.cols.get('回答納期')[r] : null;
            out.push(IGNORE_SUPPLIERS.includes(supplier) ? '対象外の発注先'
                : max !== null && isNumber(answer) && answer === max ? '★納期回答の元' : '');
            rows.push(out);
        }

        const old = wb.getWorksheet(sheetName);
        if (old) wb.removeWorksheet(old.id);
        const ds = wb.addWorksheet(sheetName);
        const widths = heads.map(() => 0);
        rows.forEach((values, ri) => {
            const xr = ds.getRow(ri + 1);
            values.forEach((v, ci) => {
                const cell = xr.getCell(ci + 1);
                if (!numberHeads.includes(heads[ci])) cell.numFmt = '@';
                cell.value = v === '' ? null : v;
                widths[ci] = Math.max(widths[ci], textWidth(v));
            });
        });
        ds.getRow(1).eachCell(cell => {
            cell.font = { bold: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD_FILL } };
        });
        ds.autoFilter = { from: { row: 1, column: 1 }, to: { row: rows.length, column: heads.length } };
        widths.forEach((w, i) => { ds.getColumn(i + 1).width = Math.min(Math.max(w + 2, 6), 60); });
        // 元のシートを開いた状態で保存する
        const index = wb.worksheets.indexOf(ws);
        if (wb.views && wb.views.length) wb.views[0].activeTab = index;
        else wb.views = [{ activeTab: index, firstSheet: 0, visibility: 'visible' }];
        for (const s of wb.worksheets) s.views = (s.views && s.views.length ? s.views : [{}]).map(v => Object.assign({}, v, { tabSelected: s === ws }));
        return pairs.length;
    }

    // 入力: zanzan / order / exclude / ledgers はどれも { name, read: () => Promise<ArrayBuffer> }、history は CSV の文字（無ければ null）
    // 出力: 書き込み済みの注残リストと、履歴・変更一覧の CSV
    async function run(opts) {
        const log = opts.log || (() => { });
        const now = opts.now || new Date();
        const nowText = formatNow(now);

        // ---- 製番台帳を読み込んで 納入指示番号 → 製番・部品納期 の対応表を作る ----
        const { map, dateMap } = await readLedgers(opts.ledgers, log);
        log(`対応表: ${map.size} 件`);
        if (map.size === 0) throw new Error('製番台帳から納入指示番号を読み取れませんでした。');

        let answers = null;
        const detail = {};
        if (opts.order) answers = await readAnswers(opts.order, log, detail);
        else log('発注仕入状況一覧表が指定されていないため、納期回答は入れません');

        let excludeItems = null;
        if (opts.exclude) excludeItems = await readExcludeItems(opts.exclude, log);
        else log('対象外品番リストが指定されていないため、品番では対象外にしません');

        // ---- 注残リストに書き込む ----
        const original = await opts.zanzan.read();
        const wb = await loadBook(original.slice(0));
        const ws = wb.worksheets[0];
        addMissingColumns(ws, log);
        const h = findHeader(ws, 5, 78, { order: /^納入指示番号$/i, seiban: /^FD製番$/i, date: /^部品納期$/i, answer: /^納期回答$/i });
        if (!h) throw new Error('注残リストに「納入指示番号」の見出しが見つかりません。');
        const first = h.row + 1;
        const last = lastRow(ws, h.order);
        if (last < first) throw new Error('注残リストにデータがありません。');

        const count = last - first + 1;
        const ex = findExcluded(ws, first, count, excludeItems, log);
        const keys = [];
        let missing = 0;
        for (let i = 0; i < count; i++) {
            keys[i] = normalizeKey(cellVal(ws.getCell(first + i, h.order)));
            if (keys[i] && !ex.flags[i] && !map.has(keys[i])) missing++;
        }

        // 対象外の行は FD製番に「対象外」と入れ、部品納期は入れない
        const s1 = fillColumn(ws, first, h.seiban, keys, map, opts.overwrite, ex.flags, EXCLUDE_TEXT);
        const s2 = fillColumn(ws, first, h.date, keys, dateMap, opts.overwrite, ex.flags, '');
        let s3 = null, s4 = null, detailRows = 0, history = null, hasRemark = false;
        if (answers) {
            history = readHistory(opts.history);
            s3 = fillAnswer(ws, first, count, h.seiban, h.answer, answers, keys, history, nowText);
            const rc = findHeader(ws, 5, 78, { remark: /^(コメント|回答備考)$/i });
            if (rc) { hasRemark = true; s4 = fillRemark(ws, first, count, h.seiban, rc.remark, answers, keys, history, nowText); }
            detailRows = writeDetailSheet(wb, ws, first, count, h.seiban, h.order, h.answer, answers, detail);
        }

        const result = {
            original,
            zanzan: await wb.xlsx.writeBuffer(),
            backupName: opts.zanzan.name.replace(/(\.[^.]+)?$/, m => `_${stamp(now)}${m}`),
            historyCsv: null,
            changeCsv: null,
            changeName: `納期回答変更_${stamp(now)}.csv`
        };

        if (s3) {
            result.historyCsv = historyCsv(history);
            if (s3.changes.length > 0) {
                const info = findHeader(ws, 5, 78, { office: /^事業所名$/i, model: /^機種名$/i, item: /^品名$/i });
                const text = (r, c) => info ? ws.getCell(r, c).text : '';
                result.changeCsv = toCsv(['区分', '行', '事業所名', '機種名', '品名', '納入指示番号', 'FD製番', '前回納期回答', '今回納期回答', '差', '前回記録日時'],
                    s3.changes.map(c => ({
                        '区分': c.kind, '行': c.row,
                        '事業所名': info && text(c.row, info.office), '機種名': info && text(c.row, info.model), '品名': info && text(c.row, info.item),
                        '納入指示番号': c.key, 'FD製番': c.seiban, '前回納期回答': c.prev, '今回納期回答': c.value,
                        '差': formatDays(c.days), '前回記録日時': c.prevTime
                    })));
            }
        }

        log('');
        log(`対象行              : ${count}`);
        log(`対象外の条件に当たる行: ${ex.rows} （ユーザ欄1がJ以外 ${ex.user} / 納入指示区分がJIT ${ex.jit} / 対象外品番 ${ex.item}。重複あり）`);
        log(`台帳に無かった行    : ${missing} （対象外の行は除く）`);
        log(`FD製番を入れた行    : ${s1.filled} （うち候補が複数で黄色: ${s1.multi}）`);
        log(`FD製番に「対象外」を入れた行: ${s1.excluded}`);
        log(`部品納期を入れた行  : ${s2.filled} （うち候補が複数で黄色: ${s2.multi}）`);
        if (s3) {
            log(`納期回答を入れた行  : ${s3.filled} （うち製番が複数で黄色: ${s3.multi} / 未回答の明細ありをメモ: ${s3.pending}）`);
            log(`  回答納期がまだ無い: ${s3.noAnswer} / 発注仕入状況一覧表に製番が無い: ${s3.notFound}`);
            const changed = s3.changes.filter(c => c.kind === '変更');
            const added = s3.changes.filter(c => c.kind === '新規');
            log(`納期回答が前回から変わった行: ${changed.length} （赤塗り・メモ付き）`);
            for (const c of changed.slice(0, 30)) log(`  ${c.key.padEnd(12)} ${c.seiban.padEnd(18)} ${c.prev} → ${c.value} ${formatDays(c.days)}`);
            if (changed.length > 30) log(`  …ほか ${changed.length - 30} 行`);
            log(`納期回答が新しく入った行    : ${added.length} （水色塗り・メモ付き）`);
            if (s3.cleared) log(`回答が無くなり納期回答を消した行: ${s3.cleared}`);
            if (hasRemark) log(`回答備考を入れた行  : ${s4.filled}${s4.cleared ? ` （備考が無くなり消した行: ${s4.cleared}）` : ''}`);
            else log('注残リストに「コメント」列が無いため、回答備考は入れていません');
            log(`「発注明細」シート  : ${detailRows} 行`);
        }
        if (!opts.overwrite) log(`入力済みで残した行  : FD製番 ${s1.kept} / 部品納期 ${s2.kept}`);
        return result;
    }

    global.Seiban = { run };
})(window);

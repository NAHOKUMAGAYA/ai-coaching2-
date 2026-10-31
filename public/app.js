// 画面の操作。Chrome / Edge では選んだファイルを覚えておき、注残リストへ上書き保存もできる
(function () {
    'use strict';

    const $ = id => document.getElementById(id);
    const canPick = 'showOpenFilePicker' in window;
    const EXCEL_ACCEPT = '.xlsx,.xlsm';
    const EXCEL_TYPES = [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'] } }];
    const CSV_TYPES = [{ description: 'CSV', accept: { 'text/csv': ['.csv'] } }];
    const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    // ファイルは { name, handle } (Chrome / Edge) か { name, file } で持つ
    const state = { zanzan: null, order: null, exclude: null, history: null, ledgers: [] };
    let result = null;

    // ---- 選んだファイルの記憶（ファイルの中身ではなく、ファイルへの参照だけを IndexedDB に置く） ----
    function openDb() {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open('seiban-tenki', 1);
            req.onupgradeneeded = () => req.result.createObjectStore('kv');
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }

    async function saveState() {
        if (!canPick) return;
        try {
            const db = await openDb();
            const handles = {
                zanzan: state.zanzan && state.zanzan.handle,
                exclude: state.exclude && state.exclude.handle,
                history: state.history && state.history.handle,
                ledgers: state.ledgers.filter(l => l.handle).map(l => l.handle)
            };
            db.transaction('kv', 'readwrite').objectStore('kv').put(handles, 'files');
        } catch (e) { /* 覚えられなくても使える */ }
    }

    async function loadState() {
        if (!canPick) return;
        try {
            const db = await openDb();
            const saved = await new Promise((resolve, reject) => {
                const req = db.transaction('kv').objectStore('kv').get('files');
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject(req.error);
            });
            if (!saved) return;
            if (saved.zanzan) state.zanzan = { name: saved.zanzan.name, handle: saved.zanzan };
            if (saved.exclude) state.exclude = { name: saved.exclude.name, handle: saved.exclude };
            if (saved.history) state.history = { name: saved.history.name, handle: saved.history };
            state.ledgers = (saved.ledgers || []).map(h => ({ name: h.name, handle: h }));
            render();
        } catch (e) { /* 前回の記憶が読めなければ選び直してもらう */ }
    }

    // ---- ファイル選択 ----
    function pickByInput(accept, multiple) {
        return new Promise(resolve => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = accept;
            input.multiple = multiple;
            input.addEventListener('change', () => resolve([...input.files].map(f => ({ name: f.name, file: f }))));
            input.click();
        });
    }

    async function pick(types, accept, multiple) {
        if (!canPick) return pickByInput(accept, multiple);
        try {
            const handles = await window.showOpenFilePicker({ types, multiple, excludeAcceptAllOption: false });
            return handles.map(h => ({ name: h.name, handle: h }));
        } catch (e) {
            if (e.name === 'AbortError') return [];
            throw e;
        }
    }

    async function ensurePermission(entry, mode) {
        if (!entry.handle) return;
        const opts = { mode };
        if (await entry.handle.queryPermission(opts) === 'granted') return;
        if (await entry.handle.requestPermission(opts) !== 'granted') throw new Error(`「${entry.name}」を開く許可がありません。`);
    }

    async function readEntry(entry) {
        return entry.handle ? entry.handle.getFile() : entry.file;
    }

    function source(entry) {
        return { name: entry.name, read: async () => (await readEntry(entry)).arrayBuffer() };
    }

    // ---- 表示 ----
    function render() {
        $('zanzanName').textContent = state.zanzan ? state.zanzan.name : '未選択';
        $('orderName').textContent = state.order ? state.order.name : '未選択';
        $('excludeName').textContent = state.exclude ? state.exclude.name : '未選択';
        $('historyName').textContent = state.history ? state.history.name : '未選択（初回）';
        const list = $('ledgerList');
        list.textContent = '';
        for (const l of state.ledgers) {
            const o = document.createElement('option');
            o.textContent = l.name;
            list.appendChild(o);
        }
    }

    const logEl = $('log');
    function log(m) {
        logEl.textContent += m + '\n';
        logEl.scrollTop = logEl.scrollHeight;
    }

    function addLedgers(entries) {
        for (const e of entries) {
            if (!/\.(xlsx|xlsm)$/i.test(e.name)) {
                if (/\.xls$/i.test(e.name)) log(`※ ${e.name} は古い形式(.xls)なので読めません。Excel で .xlsx に保存し直してください。`);
                continue;
            }
            if (!state.ledgers.some(l => l.name === e.name)) state.ledgers.push(e);
        }
        render();
        saveState();
    }

    $('zanzanPick').addEventListener('click', async () => {
        const [e] = await pick(EXCEL_TYPES, EXCEL_ACCEPT, false);
        if (!e) return;
        state.zanzan = e;
        render();
        saveState();
    });

    $('orderPick').addEventListener('click', async () => {
        const [e] = await pick(EXCEL_TYPES, EXCEL_ACCEPT, false);
        if (!e) return;
        state.order = e;
        render();
    });
    $('orderClear').addEventListener('click', () => { state.order = null; render(); });

    $('excludePick').addEventListener('click', async () => {
        const [e] = await pick(EXCEL_TYPES, EXCEL_ACCEPT, false);
        if (!e) return;
        state.exclude = e;
        render();
        saveState();
    });
    $('excludeClear').addEventListener('click', () => { state.exclude = null; render(); saveState(); });

    $('historyPick').addEventListener('click', async () => {
        const [e] = await pick(CSV_TYPES, '.csv', false);
        if (!e) return;
        state.history = e;
        render();
        saveState();
    });
    $('historyClear').addEventListener('click', () => { state.history = null; render(); saveState(); });

    $('ledgerAdd').addEventListener('click', async () => addLedgers(await pick(EXCEL_TYPES, EXCEL_ACCEPT, true)));
    $('ledgerDel').addEventListener('click', () => {
        const selected = [...$('ledgerList').options].map(o => o.selected);
        state.ledgers = state.ledgers.filter((l, i) => !selected[i]);
        render();
        saveState();
    });

    // エクスプローラーから台帳ファイルをドラッグして追加できるようにする
    const list = $('ledgerList');
    list.addEventListener('dragover', e => { e.preventDefault(); list.classList.add('drag'); });
    list.addEventListener('dragleave', () => list.classList.remove('drag'));
    list.addEventListener('drop', async e => {
        e.preventDefault();
        list.classList.remove('drag');
        // ハンドルはイベントの中で同期的に取り出しておく必要がある
        const items = [...e.dataTransfer.items].filter(i => i.kind === 'file');
        const pending = items.map(i => i.getAsFileSystemHandle ? i.getAsFileSystemHandle() : Promise.resolve(null));
        const files = items.map(i => i.getAsFile());
        const handles = await Promise.all(pending);
        addLedgers(files.map((f, i) => handles[i] && handles[i].kind === 'file' ? { name: handles[i].name, handle: handles[i] } : { name: f.name, file: f }));
    });

    // 選んだファイル・チェック・処理の記録をすべて消して、初めて開いた状態に戻す
    $('clear').addEventListener('click', () => {
        if (!confirm('選んだファイルと処理の記録をすべて消します。よろしいですか？')) return;
        state.zanzan = null;
        state.order = null;
        state.exclude = null;
        state.history = null;
        state.ledgers = [];
        result = null;
        $('overwrite').checked = false;
        $('result').hidden = true;
        logEl.textContent = '';
        render();
        saveState();
    });

    // ---- 実行 ----
    $('run').addEventListener('click', async () => {
        const btn = $('run');
        btn.disabled = true;
        $('clear').disabled = true;
        $('result').hidden = true;
        result = null;
        logEl.textContent = '';
        try {
            if (typeof ExcelJS === 'undefined') throw new Error('Excel を読むための部品(ExcelJS)を読み込めませんでした。ネットにつながっているか確認して、ページを開き直してください。');
            if (!state.zanzan) throw new Error('注残リストを選んでください。');
            if (!/\.xlsx$/i.test(state.zanzan.name)) throw new Error('注残リストは .xlsx のファイルを選んでください。');
            if (state.ledgers.length === 0) throw new Error('製番台帳を1つ以上追加してください。');
            // 前回のファイルを使うときは、ここでまとめて読む許可をもらう
            for (const e of [state.zanzan, state.order, state.exclude, state.history, ...state.ledgers]) {
                if (e) await ensurePermission(e, 'read');
            }
            const historyText = state.history ? await (await readEntry(state.history)).text() : null;
            result = await Seiban.run({
                zanzan: source(state.zanzan),
                order: state.order ? source(state.order) : null,
                exclude: state.exclude ? source(state.exclude) : null,
                ledgers: state.ledgers.map(source),
                history: historyText,
                overwrite: $('overwrite').checked,
                log
            });
            showResult();
        } catch (e) {
            log(`エラー: ${e.message}`);
            alert(e.message);
        } finally {
            btn.disabled = false;
            $('clear').disabled = false;
        }
    });

    function showResult() {
        const over = canPick && state.zanzan.handle;
        $('saveOver').hidden = !over;
        $('saveOver').disabled = false;
        const parts = [];
        if (over) parts.push(`「注残リストに上書き保存」は ${state.zanzan.name} を書き換えます。書き換える前の元のファイルは ${result.backupName} としてダウンロードフォルダに残します。`);
        parts.push(`「ダウンロードで保存」は ${state.zanzan.name} をダウンロードフォルダに保存します。`);
        if (result.historyCsv) parts.push(state.history && state.history.handle && over
            ? `納期回答履歴は選んだ ${state.history.name} に上書きします（ダウンロードで保存のときはダウンロードします）。`
            : '納期回答履歴.csv はダウンロードされるので、注残リストと同じフォルダに置いて次回選んでください。');
        if (result.changeCsv) parts.push(`納期回答の変更・新規の一覧 ${result.changeName} もダウンロードします。`);
        $('saveNote').textContent = parts.join(' ');
        $('result').hidden = false;
        log('');
        log('処理が終わりました。下の「保存」から保存してください（まだ保存されていません）。');
    }

    function download(name, data, type) {
        const url = URL.createObjectURL(new Blob([data], { type }));
        const a = document.createElement('a');
        a.href = url;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 60000);
    }

    async function writeHandle(entry, data) {
        await ensurePermission(entry, 'readwrite');
        const w = await entry.handle.createWritable();
        try {
            await w.write(data);
            await w.close();
        } catch (e) {
            await w.abort().catch(() => { });
            throw new Error(`「${entry.name}」に保存できませんでした。Excel などで開いていたら閉じてから、もう一度保存してください。（${e.message}）`);
        }
    }

    function downloadCsvs(includeHistory) {
        if (includeHistory && result.historyCsv) download('納期回答履歴.csv', result.historyCsv, 'text/csv');
        if (result.changeCsv) download(result.changeName, result.changeCsv, 'text/csv');
    }

    $('saveOver').addEventListener('click', async () => {
        try {
            // 上書きの許可を先にもらってから、元のファイルを控えておく
            await ensurePermission(state.zanzan, 'readwrite');
            const historyOver = result.historyCsv && state.history && state.history.handle;
            if (historyOver) await ensurePermission(state.history, 'readwrite');
            download(result.backupName, result.original, XLSX_MIME);
            await writeHandle(state.zanzan, result.zanzan);
            log(`保存しました: ${state.zanzan.name}（元のファイルは ${result.backupName} としてダウンロード）`);
            // 注残リストを保存できてから、履歴と変更一覧を書き出す
            if (historyOver) {
                await writeHandle(state.history, result.historyCsv);
                log(`納期回答履歴を更新しました: ${state.history.name}`);
            }
            downloadCsvs(!historyOver);
            $('saveOver').disabled = true;
        } catch (e) {
            log(`エラー: ${e.message}`);
            alert(e.message);
        }
    });

    $('saveDownload').addEventListener('click', () => {
        download(state.zanzan.name, result.zanzan, XLSX_MIME);
        downloadCsvs(true);
        log(`ダウンロードしました: ${state.zanzan.name}`);
    });

    $('support').hidden = canPick;
    render();
    loadState();
})();

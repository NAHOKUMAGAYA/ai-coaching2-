// 画面の操作。Chrome / Edge では選んだファイルを覚えておき、選んだ注残リストへ上書き保存し、履歴・変更一覧も同じフォルダに保存する
(function () {
    'use strict';

    const $ = id => document.getElementById(id);
    const canPick = 'showOpenFilePicker' in window;
    const EXCEL_ACCEPT = '.xlsx,.xlsm';
    const EXCEL_TYPES = [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'] } }];
    const CSV_TYPES = [{ description: 'CSV', accept: { 'text/csv': ['.csv'] } }];
    const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    // 注残リストのフォルダに上書き保存できるか（Chrome / Edge）。できなければダウンロードで保存する
    const canSaveFolder = 'showDirectoryPicker' in window;
    const BACKUP_DIR = 'バックアップ';

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

    // Edge などで、ファイルを選んでも「キャンセルされた」(AbortError) と返されることがある。
    // そうなったら、次からは昔ながらのファイル選択(input)を使う（この場合も、保存のときに注残リストのフォルダを選べば上書きできる）
    const PICK_KEY = 'seiban-tenki-pick';
    function useInputPicker() {
        try { return localStorage.getItem(PICK_KEY) === 'input'; } catch (e) { return false; }
    }

    async function pick(types, accept, multiple) {
        if (!canPick || useInputPicker()) return pickByInput(accept, multiple);
        try {
            const handles = await window.showOpenFilePicker({ types, multiple, excludeAcceptAllOption: false });
            if (handles.length === 0) log('※ファイル選択: ファイルが返ってきませんでした');
            return handles.map(h => ({ name: h.name, handle: h }));
        } catch (e) {
            // キャンセルでも、ブラウザに止められたときも AbortError になることがあるので、内容を下の欄に残す
            if (e.name === 'AbortError') {
                try { localStorage.setItem(PICK_KEY, 'input'); } catch (err) { /* 覚えられなくても次の操作はできる */ }
                log('※ファイルが選ばれませんでした。選んだのに欄が変わらなかったときは、もう一度「選ぶ...」を押してください（別の方法で開きます）。');
                return [];
            }
            // 選べなかった理由が分かるよう、黙って止まらずに知らせる
            const msg = `ファイルを選べませんでした。\n（${e.name}: ${e.message}）\n\n` +
                'メールやTeamsから直接開いたファイルや、システムのフォルダにあるファイルは選べないことがあります。' +
                'デスクトップなどに保存し直してから選んでください。';
            log(msg);
            alert(msg);
            return [];
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
        $('saveOver').hidden = !canSaveFolder;
        $('saveOver').disabled = false;
        $('saveDownloadNote').hidden = !canSaveFolder;
        const parts = [];
        if (canSaveFolder) {
            parts.push(`「保存」を押すと、選んだ注残リスト ${state.zanzan.name} を上書きします。` +
                `納期回答履歴・変更一覧も同じフォルダに保存し、書き換える前の注残リストは同じフォルダの「${BACKUP_DIR}」フォルダに ${result.backupName} として残します。` +
                '初めて保存するときは、注残リストのあるフォルダを選ぶ画面が出ます。');
        } else {
            parts.push(`このブラウザでは上書き保存ができないので、${state.zanzan.name} と納期回答履歴・変更一覧をダウンロードフォルダに保存します。注残リストと同じフォルダへ移してください。`);
        }
        $('saveNote').textContent = parts.join(' ');
        $('result').hidden = false;
        log('');
        log('処理が終わりました。下の「保存」を押してください（まだ保存されていません）。');
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

    async function writeFile(handle, data) {
        const w = await handle.createWritable();
        try {
            await w.write(data);
            await w.close();
        } catch (e) {
            await w.abort().catch(() => { });
            throw new Error(`「${handle.name}」に保存できませんでした。Excel などで開いていたら閉じてから、もう一度「保存」を押してください。（${e.message}）`);
        }
    }

    // ---- 保存先のフォルダ（注残リストのあるフォルダ）----
    // ファイル1つの参照からはフォルダが分からないので、フォルダの参照をもらって覚えておく

    async function loadFolder() {
        try {
            const db = await openDb();
            return await new Promise((resolve, reject) => {
                const req = db.transaction('kv').objectStore('kv').get('folder');
                req.onsuccess = () => resolve(req.result || null);
                req.onerror = () => reject(req.error);
            });
        } catch (e) { return null; }
    }

    async function saveFolder(dir) {
        try {
            const db = await openDb();
            db.transaction('kv', 'readwrite').objectStore('kv').put(dir, 'folder');
        } catch (e) { /* 覚えられなければ次も選んでもらう */ }
    }

    // そのフォルダに、選んだ注残リストそのものがあればその参照を返す（無ければ null）
    async function findZanzanIn(dir) {
        let h;
        try { h = await dir.getFileHandle(state.zanzan.name); } catch (e) { return null; }
        if (state.zanzan.handle) return await state.zanzan.handle.isSameEntry(h) ? h : null;
        // 昔ながらの選び方のときは、名前・大きさ・更新日時が同じなら同じファイルとみなす
        const f = await h.getFile();
        const sel = state.zanzan.file;
        return f.size === sel.size && f.lastModified === sel.lastModified ? h : null;
    }

    async function hasPermission(dir) {
        const opts = { mode: 'readwrite' };
        if (await dir.queryPermission(opts) === 'granted') return true;
        return await dir.requestPermission(opts) === 'granted';
    }

    // 前回のフォルダに注残リストがあればそれを使い、無ければフォルダを選んでもらう
    async function getFolder() {
        const saved = await loadFolder();
        if (saved && await hasPermission(saved).catch(() => false)) {
            const h = await findZanzanIn(saved).catch(() => null);
            if (h) return { dir: saved, zanzan: h };
        }
        const opts = { id: 'seiban-tenki-folder', mode: 'readwrite' };
        if (state.zanzan.handle) opts.startIn = state.zanzan.handle;
        let dir;
        try {
            dir = await window.showDirectoryPicker(opts);
        } catch (e) {
            if (e.name === 'AbortError') return null;
            throw e;
        }
        const h = await findZanzanIn(dir);
        if (!h) throw new Error(`選んだフォルダに、選んだ注残リスト「${state.zanzan.name}」がありません。注残リストのあるフォルダを選んでください。`);
        await saveFolder(dir);
        return { dir, zanzan: h };
    }

    // 納期回答履歴の保存先。選んだ履歴があればそれに上書きし、無ければ注残リストのフォルダに作る
    async function historyTarget(dir) {
        if (state.history && state.history.handle) {
            await ensurePermission(state.history, 'readwrite');
            return { handle: state.history.handle };
        }
        if (state.history) return { handle: await dir.getFileHandle(state.history.name, { create: true }) };
        const name = '納期回答履歴.csv';
        let exists = true;
        try { await dir.getFileHandle(name); } catch (e) { exists = false; }
        if (!exists) return { handle: await dir.getFileHandle(name, { create: true }) };
        // 今回読まなかった履歴を上書きすると過去の分が消えるので、別の名前で残す
        const other = name.replace(/\.csv$/, `_${result.changeName.match(/_(\d+_\d+)\.csv$/)[1]}.csv`);
        return {
            handle: await dir.getFileHandle(other, { create: true }),
            warn: `※フォルダに ${name} がありますが、今回は「納期回答履歴.csv」欄で選ばれていなかったので上書きせず、${other} として保存しました。次回は ${name} を選んでから実行してください。`
        };
    }

    $('saveOver').addEventListener('click', async () => {
        const btn = $('saveOver');
        btn.disabled = true;
        try {
            const folder = await getFolder();
            if (!folder) { log('※保存先のフォルダが選ばれなかったので、保存していません。'); btn.disabled = false; return; }
            const { dir, zanzan } = folder;
            const history = result.historyCsv ? await historyTarget(dir) : null;

            // 書き換える前の注残リストを控えてから上書きする
            const backupDir = await dir.getDirectoryHandle(BACKUP_DIR, { create: true });
            await writeFile(await backupDir.getFileHandle(result.backupName, { create: true }), result.original);
            await writeFile(zanzan, result.zanzan);
            log(`保存しました: ${dir.name}\\${state.zanzan.name}（元のファイルは ${BACKUP_DIR}\\${result.backupName}）`);

            // 注残リストを保存できてから、履歴と変更一覧を書き出す
            if (history) {
                await writeFile(history.handle, result.historyCsv);
                log(`納期回答履歴を保存しました: ${history.handle.name}`);
                if (history.warn) log(history.warn);
            }
            if (result.changeCsv) {
                await writeFile(await dir.getFileHandle(result.changeName, { create: true }), result.changeCsv);
                log(`納期回答の変更一覧を保存しました: ${result.changeName}`);
            }
            log('');
            log('保存が終わりました。');
        } catch (e) {
            btn.disabled = false;
            log(`エラー: ${e.message}`);
            alert(e.message);
        }
    });

    $('saveDownload').addEventListener('click', () => {
        download(state.zanzan.name, result.zanzan, XLSX_MIME);
        if (result.historyCsv) download('納期回答履歴.csv', result.historyCsv, 'text/csv');
        if (result.changeCsv) download(result.changeName, result.changeCsv, 'text/csv');
        log(`ダウンロードフォルダに保存しました: ${state.zanzan.name}（注残リストと同じフォルダへ移してください）`);
    });

    $('support').hidden = canSaveFolder;
    render();
    loadState();
})();

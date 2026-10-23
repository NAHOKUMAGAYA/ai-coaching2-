# 注残リストの「納入指示番号」から製番台帳を引いて「FD製番」「部品納期」列に書き込む
# 使い方: 製番転記.bat をダブルクリック
#   コマンドで実行する場合: powershell -File 製番転記.ps1 -NoGui [-Zanzan パス] [-Overwrite]
param(
    [string]$Zanzan = 'C:\Users\n-kumagaya\Desktop\ai-coaching2-\納期回答関係\注残リスト.xlsx',
    [string[]]$Ledgers = @(
        '\\192.168.11.200\Share\12.製番台帳\43期製番台帳\（新）4374(AMD単品(富士宮)).xlsx',
        '\\192.168.11.200\Share\12.製番台帳\43期製番台帳\（新）4366(AMD単品(土岐)).xlsx'
    ),
    [switch]$Overwrite,
    [switch]$NoGui
)

$ErrorActionPreference = 'Stop'
$SkipSheets = @('記入方法', 'リスト', '休日')
$MultiColor = 65535   # 黄色: 候補が複数ある行

function Normalize-Key([string]$s) {
    if (-not $s) { return '' }
    return $s.Normalize([Text.NormalizationForm]::FormKC).Trim().ToUpper()
}

# 2次元配列の 1～maxRow 行目から、指定パターンの見出しがすべて揃う行を探す
function Find-Header($v, [int]$maxRow, [hashtable]$patterns) {
    $rows = [Math]::Min($maxRow, $v.GetUpperBound(0))
    $cols = $v.GetUpperBound(1)
    for ($r = 1; $r -le $rows; $r++) {
        $found = @{}
        for ($c = 1; $c -le $cols; $c++) {
            $t = ("$($v[$r, $c])" -replace '\s', '')
            foreach ($name in $patterns.Keys) {
                if (-not $found.ContainsKey($name) -and $t -match $patterns[$name]) { $found[$name] = $c }
            }
        }
        if ($found.Count -eq $patterns.Count) { $found['row'] = $r; return $found }
    }
    return $null
}

function Release($o) {
    if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
}

function Add-Value([hashtable]$map, [string]$key, [string]$value) {
    if (-not $value) { return }
    if (-not $map.ContainsKey($key)) { $map[$key] = New-Object System.Collections.Generic.List[string] }
    if (-not $map[$key].Contains($value)) { $map[$key].Add($value) }
}

# 台帳の日付(シリアル値)を注残リストの他の日付列と同じ yyyy-MM-dd の文字にする
function Format-Date($v) {
    if ($v -is [double]) { return [DateTime]::FromOADate($v).ToString('yyyy-MM-dd') }
    return "$v".Trim()
}

# 注残リストの1列を、納入指示番号をキーに対応表から埋める。候補が複数ある行は「/」でつないで黄色にする
function Fill-Column($ws, [int]$first, [int]$col, [string[]]$keys, [hashtable]$lookup, [bool]$overwrite) {
    $count = $keys.Count
    $range = $ws.Range($ws.Cells.Item($first, $col), $ws.Cells.Item($first + $count - 1, $col))
    $current = $range.Value2
    $out = New-Object 'object[,]' $count, 1
    $multiRows = New-Object System.Collections.Generic.List[int]
    $stat = @{ filled = 0; multi = 0; kept = 0 }

    for ($i = 0; $i -lt $count; $i++) {
        $existing = $current[($i + 1), 1]
        $out[$i, 0] = $existing
        if (-not $overwrite -and "$existing" -ne '') { $stat.kept++; continue }
        $k = $keys[$i]
        if (-not $k -or -not $lookup.ContainsKey($k)) { continue }
        $out[$i, 0] = ($lookup[$k] -join '/')
        $stat.filled++
        if ($lookup[$k].Count -gt 1) { $stat.multi++; $multiRows.Add($first + $i) }
    }

    # 製番や日付が数値に変換されないよう文字列書式にしてから一括書き込み
    $range.NumberFormat = '@'
    $range.Value2 = $out
    foreach ($r in $multiRows) { $ws.Cells.Item($r, $col).Interior.Color = $MultiColor }
    Release $range
    return $stat
}

function Invoke-Fill([string]$zanzanPath, [string[]]$ledgerPaths, [bool]$overwrite, [scriptblock]$log) {
    if (-not (Test-Path -LiteralPath $zanzanPath)) { throw "注残リストが見つかりません: $zanzanPath" }

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    try {
        # ---- 製番台帳を読み込んで 納入指示番号 → 製番・部品納期 の対応表を作る ----
        $map = @{}
        $dateMap = @{}
        foreach ($path in $ledgerPaths) {
            & $log "台帳読込: $(Split-Path $path -Leaf)"
            if (-not (Test-Path -LiteralPath $path)) { & $log "  ※ファイルが見つかりません。スキップします"; continue }
            $wb = $xl.Workbooks.Open($path, 0, $true)
            try {
                foreach ($ws in $wb.Worksheets) {
                    if ($SkipSheets -contains $ws.Name) { continue }
                    $head = $ws.Range('A1:AJ10').Value2
                    $h = Find-Header $head 10 @{ seiban = '^製番$'; order = '客先O/N' }
                    if (-not $h) { continue }
                    $last = $ws.Cells.Item($ws.Rows.Count, $h.seiban).End(-4162).Row
                    if ($last -le $h.row) { continue }
                    $sb = $ws.Range($ws.Cells.Item(1, $h.seiban), $ws.Cells.Item($last, $h.seiban)).Value2
                    $on = $ws.Range($ws.Cells.Item(1, $h.order), $ws.Cells.Item($last, $h.order)).Value2
                    $dateCol = 0
                    for ($c = 1; $c -le $head.GetUpperBound(1); $c++) {
                        if (("$($head[$h.row, $c])" -replace '\s', '') -eq '部品納期') { $dateCol = $c; break }
                    }
                    # if式の結果に代入すると2次元配列が平たくなるので、中で代入する
                    $dt = $null
                    if ($dateCol) { $dt = $ws.Range($ws.Cells.Item(1, $dateCol), $ws.Cells.Item($last, $dateCol)).Value2 }
                    $n = 0
                    for ($r = $h.row + 1; $r -le $last; $r++) {
                        $seiban = "$($sb[$r, 1])".Trim()
                        $orders = "$($on[$r, 1])"
                        if (-not $seiban -or -not $orders) { continue }
                        $date = if ($dateCol) { Format-Date $dt[$r, 1] } else { '' }
                        foreach ($k in ($orders -split '[,、，;；\s]+')) {
                            $k = Normalize-Key $k
                            if (-not $k) { continue }
                            Add-Value $map $k $seiban
                            Add-Value $dateMap $k $date
                            $n++
                        }
                    }
                    $note = if ($dateCol) { '' } else { '（部品納期の列なし）' }
                    & $log "  シート[$($ws.Name)] 納入指示番号 $n 件$note"
                }
            } finally {
                $wb.Close($false)
                Release $wb
            }
        }
        & $log "対応表: $($map.Count) 件"
        if ($map.Count -eq 0) { throw '製番台帳から納入指示番号を読み取れませんでした。' }

        # ---- 注残リストに書き込む ----
        $wb = $xl.Workbooks.Open($zanzanPath, 0, $false)
        try {
            if ($wb.ReadOnly) { throw '注残リストが他で開かれています。Excelで閉じてから実行してください。' }
            $ws = $wb.Worksheets.Item(1)
            $h = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ order = '^納入指示番号$'; seiban = '^FD製番$'; date = '^部品納期$' }
            if (-not $h) { throw '注残リストに「納入指示番号」「FD製番」「部品納期」の見出しが見つかりません。' }
            $first = $h.row + 1
            $last = $ws.Cells.Item($ws.Rows.Count, $h.order).End(-4162).Row
            if ($last -lt $first) { throw '注残リストにデータがありません。' }

            $orders = $ws.Range($ws.Cells.Item($first, $h.order), $ws.Cells.Item($last, $h.order)).Value2
            $count = $last - $first + 1
            $keys = New-Object 'string[]' $count
            $missing = 0
            for ($i = 0; $i -lt $count; $i++) {
                $keys[$i] = Normalize-Key "$($orders[($i + 1), 1])"
                if ($keys[$i] -and -not $map.ContainsKey($keys[$i])) { $missing++ }
            }

            $s1 = Fill-Column $ws $first $h.seiban $keys $map $overwrite
            $s2 = Fill-Column $ws $first $h.date $keys $dateMap $overwrite

            # 上書き前の状態をバックアップしてから保存
            $dir = Join-Path (Split-Path $zanzanPath) 'バックアップ'
            if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
            $backup = Join-Path $dir ("{0}_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($zanzanPath), (Get-Date -Format 'yyyyMMdd_HHmmss'), [IO.Path]::GetExtension($zanzanPath))
            Copy-Item -LiteralPath $zanzanPath -Destination $backup
            $wb.Save()

            & $log ''
            & $log "対象行              : $count"
            & $log "台帳に無かった行    : $missing"
            & $log "FD製番を入れた行    : $($s1.filled) （うち候補が複数で黄色表示: $($s1.multi)）"
            & $log "部品納期を入れた行  : $($s2.filled) （うち候補が複数で黄色表示: $($s2.multi)）"
            if (-not $overwrite) { & $log "入力済みで残した行  : FD製番 $($s1.kept) / 部品納期 $($s2.kept)" }
            & $log "保存しました。元のファイルは $backup に残しています。"
        } finally {
            $wb.Close($false)
            Release $ws
            Release $wb
        }
    } finally {
        $xl.Quit()
        Release $xl
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if ($NoGui) {
    Invoke-Fill $Zanzan $Ledgers $Overwrite.IsPresent { param($m) Write-Host $m }
    return
}

# ---------------- 画面 ----------------
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()

$font = New-Object Drawing.Font('Meiryo UI', 9)
$form = New-Object Windows.Forms.Form
$form.Text = '注残リスト 製番転記'
$form.Size = New-Object Drawing.Size(760, 600)
$form.StartPosition = 'CenterScreen'
$form.Font = $font

$lbl1 = New-Object Windows.Forms.Label
$lbl1.Text = '注残リスト'
$lbl1.Location = New-Object Drawing.Point(12, 15); $lbl1.AutoSize = $true
$txtZ = New-Object Windows.Forms.TextBox
$txtZ.Text = $Zanzan
$txtZ.Location = New-Object Drawing.Point(12, 35); $txtZ.Size = New-Object Drawing.Size(620, 24)
$txtZ.Anchor = 'Top,Left,Right'
$btnZ = New-Object Windows.Forms.Button
$btnZ.Text = '参照...'
$btnZ.Location = New-Object Drawing.Point(640, 33); $btnZ.Size = New-Object Drawing.Size(90, 27)
$btnZ.Anchor = 'Top,Right'
$btnZ.Add_Click({
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel|*.xlsx;*.xlsm;*.xls'
    if ($dlg.ShowDialog($form) -eq 'OK') { $txtZ.Text = $dlg.FileName }
})

$lbl2 = New-Object Windows.Forms.Label
$lbl2.Text = '製番台帳（複数可・ここへファイルをドラッグしても追加できます）'
$lbl2.Location = New-Object Drawing.Point(12, 72); $lbl2.AutoSize = $true
$lst = New-Object Windows.Forms.ListBox
$lst.Location = New-Object Drawing.Point(12, 92); $lst.Size = New-Object Drawing.Size(620, 100)
$lst.Anchor = 'Top,Left,Right'
$lst.HorizontalScrollbar = $true
$lst.SelectionMode = 'MultiExtended'
foreach ($l in $Ledgers) { [void]$lst.Items.Add($l) }

function Add-Ledgers([string[]]$paths) {
    foreach ($f in $paths) {
        if ($f -match '\.(xlsx|xlsm|xls)$' -and -not $lst.Items.Contains($f)) { [void]$lst.Items.Add($f) }
    }
}

# エクスプローラーから台帳ファイルをドラッグして追加できるようにする
$lst.AllowDrop = $true
$lst.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = 'Copy' }
})
$lst.Add_DragDrop({ Add-Ledgers $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop) })

$btnAdd = New-Object Windows.Forms.Button
$btnAdd.Text = '追加...'
$btnAdd.Location = New-Object Drawing.Point(640, 92); $btnAdd.Size = New-Object Drawing.Size(90, 27)
$btnAdd.Anchor = 'Top,Right'
$btnAdd.Add_Click({
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel|*.xlsx;*.xlsm;*.xls'
    $dlg.Multiselect = $true
    if ($lst.Items.Count -gt 0) { $dlg.InitialDirectory = [IO.Path]::GetDirectoryName($lst.Items[$lst.Items.Count - 1]) }
    # 親ウィンドウを指定しないと、2回目以降に選択画面がアプリの裏に隠れることがある
    if ($dlg.ShowDialog($form) -eq 'OK') { Add-Ledgers $dlg.FileNames }
    $dlg.Dispose()
})
$btnDel = New-Object Windows.Forms.Button
$btnDel.Text = '削除'
$btnDel.Location = New-Object Drawing.Point(640, 125); $btnDel.Size = New-Object Drawing.Size(90, 27)
$btnDel.Anchor = 'Top,Right'
$btnDel.Add_Click({
    foreach ($item in @($lst.SelectedItems)) { $lst.Items.Remove($item) }
})

$chk = New-Object Windows.Forms.CheckBox
$chk.Text = 'すでにFD製番・部品納期が入っている行も上書きする'
$chk.Checked = $Overwrite.IsPresent
$chk.Location = New-Object Drawing.Point(12, 202); $chk.AutoSize = $true

$btnRun = New-Object Windows.Forms.Button
$btnRun.Text = '実行'
$btnRun.Location = New-Object Drawing.Point(640, 198); $btnRun.Size = New-Object Drawing.Size(90, 32)
$btnRun.Anchor = 'Top,Right'

$txtLog = New-Object Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Location = New-Object Drawing.Point(12, 240); $txtLog.Size = New-Object Drawing.Size(718, 305)
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$txtLog.Font = New-Object Drawing.Font('MS Gothic', 9)

$log = {
    param($m)
    $txtLog.AppendText("$m`r`n")
    [Windows.Forms.Application]::DoEvents()
}

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $form.Cursor = 'WaitCursor'
    $txtLog.Clear()
    try {
        Invoke-Fill $txtZ.Text @($lst.Items) $chk.Checked $log
        [Windows.Forms.MessageBox]::Show($form, '完了しました。', '製番転記') | Out-Null
    } catch {
        & $log "エラー: $($_.Exception.Message)"
        [Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'エラー', 'OK', 'Error') | Out-Null
    } finally {
        $form.Cursor = 'Default'
        $btnRun.Enabled = $true
    }
})

$form.Controls.AddRange(@($lbl1, $txtZ, $btnZ, $lbl2, $lst, $btnAdd, $btnDel, $chk, $btnRun, $txtLog))
[void]$form.ShowDialog()

# 注残リストの「納入指示番号」から製番台帳を引いて「FD製番」「部品納期」列に書き込み、
# その製番の発注明細(発注仕入状況一覧表)から一番遅い回答納期を「納期回答」列に書き込む
# 使い方: 製番転記.bat をダブルクリック
#   コマンドで実行する場合: powershell -File 製番転記.ps1 -NoGui [-Zanzan パス] [-Order パス] [-Overwrite]
param(
    [string]$Zanzan = 'C:\Users\n-kumagaya\Desktop\ai-coaching2-\納期回答関係\注残リスト.xlsx',
    [string]$Order = '',   # 発注仕入状況一覧表。省略時は注残リストと同じフォルダの一番新しいもの
    [string[]]$Ledgers = @(
        '\\192.168.11.200\Share\12.製番台帳\43期製番台帳\（新）4374(AMD単品(富士宮)).xlsx',
        '\\192.168.11.200\Share\12.製番台帳\43期製番台帳\（新）4366(AMD単品(土岐)).xlsx'
    ),
    [switch]$Overwrite,
    [switch]$NoGui
)

$ErrorActionPreference = 'Stop'
$SkipSheets = @('記入方法', 'リスト', '休日')

# 画面で選んだ注残リストと製番台帳を覚えておき、次回もそのまま使う
$SettingsPath = Join-Path $PSScriptRoot '製番転記_設定.json'
if (Test-Path -LiteralPath $SettingsPath) {
    $saved = Get-Content -LiteralPath $SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($saved.Zanzan -and -not $PSBoundParameters.ContainsKey('Zanzan')) { $Zanzan = $saved.Zanzan }
    if ($saved.Ledgers -and -not $PSBoundParameters.ContainsKey('Ledgers')) { $Ledgers = @($saved.Ledgers) }
}

function Save-Settings([string]$zanzanPath, [string[]]$ledgerPaths) {
    ConvertTo-Json -InputObject ([ordered]@{ Zanzan = $zanzanPath; Ledgers = @($ledgerPaths) }) |
        Set-Content -LiteralPath $SettingsPath -Encoding UTF8
}
$MultiColor = 65535   # 黄色: 候補が複数ある行
$PendingColor = 49407 # オレンジ: 回答納期が未入力の発注明細が残っている行
$ChangedColor = 255      # 赤塗り: 前回から納期回答が変わった行
$AddedColor = 15773696   # 水色塗り: 空欄だった納期回答が新しく入った行
$MarkFontColor = 16777215 # 赤塗り・水色塗りの文字は白

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

function Find-OrderFile([string]$zanzanPath) {
    $dir = Split-Path $zanzanPath
    if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { return '' }
    $f = Get-ChildItem -LiteralPath $dir -Filter '発注仕入状況一覧表*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($f) { return $f.FullName }
    return ''
}

# 発注仕入状況一覧表から 製造番号 → 一番遅い回答納期・未回答の明細の有無 を集計する
function Read-Answers($xl, [string]$path, [scriptblock]$log) {
    & $log "発注仕入状況一覧表読込: $(Split-Path $path -Leaf)"
    $answers = @{}
    $wb = $xl.Workbooks.Open($path, 0, $true)
    try {
        $ws = $wb.Worksheets.Item(1)
        $h = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ seiban = '^製造番号$'; answer = '^回答納期$'; remain = '^仕入残$' }
        if (-not $h) { throw '発注仕入状況一覧表に「製造番号」「回答納期」「仕入残」の見出しが見つかりません。' }
        $last = $ws.Cells.Item($ws.Rows.Count, $h.seiban).End(-4162).Row
        $sb = $ws.Range($ws.Cells.Item(1, $h.seiban), $ws.Cells.Item($last, $h.seiban)).Value2
        $an = $ws.Range($ws.Cells.Item(1, $h.answer), $ws.Cells.Item($last, $h.answer)).Value2
        $rm = $ws.Range($ws.Cells.Item(1, $h.remain), $ws.Cells.Item($last, $h.remain)).Value2
        for ($r = $h.row + 1; $r -le $last; $r++) {
            $k = Normalize-Key "$($sb[$r, 1])"
            if (-not $k) { continue }
            if (-not $answers.ContainsKey($k)) { $answers[$k] = @{ max = $null; pending = $false } }
            $e = $answers[$k]
            $a = $an[$r, 1]
            if ($a -is [double]) {
                if ($null -eq $e.max -or $a -gt $e.max) { $e.max = $a }
            } elseif ($rm[$r, 1] -is [double] -and $rm[$r, 1] -gt 0) {
                # 入荷が残っているのに回答納期が空の明細
                $e.pending = $true
            }
        }
        & $log "  発注明細 $($last - $h.row) 行 / 製造番号 $($answers.Count) 件"
    } finally {
        $wb.Close($false)
        Release $wb
    }
    return $answers
}

# 前回までに入れた納期回答の記録（納入指示番号ごと）
function Read-History([string]$path) {
    $history = @{}
    if (Test-Path -LiteralPath $path) {
        foreach ($row in (Import-Csv -LiteralPath $path -Encoding UTF8)) {
            $history[$row.'納入指示番号'] = @{ value = $row.'納期回答'; seiban = $row.'FD製番'; time = $row.'記録日時' }
        }
    }
    return $history
}

function Save-History([string]$path, [hashtable]$history) {
    $history.GetEnumerator() | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ '納入指示番号' = $_.Name; 'FD製番' = $_.Value.seiban; '納期回答' = $_.Value.value; '記録日時' = $_.Value.time }
    } | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
}

# FD製番ごとに一番遅い回答納期を「納期回答」列に毎回入れ直す。FD製番が「/」で複数ならその中で一番遅い日。
# 前回の記録(無ければセルにあった値)と変わった行は赤塗り、空欄から入った行は水色塗りにして、内容をメモに残す
function Fill-Answer($ws, [int]$first, [int]$count, [int]$seibanCol, [int]$col, [hashtable]$answers, [string[]]$keys, [hashtable]$history) {
    $seibans = $ws.Range($ws.Cells.Item($first, $seibanCol), $ws.Cells.Item($first + $count - 1, $seibanCol)).Value2
    $range = $ws.Range($ws.Cells.Item($first, $col), $ws.Cells.Item($first + $count - 1, $col))
    $current = $range.Value2
    $out = New-Object 'object[,]' $count, 1
    $pendingRows = New-Object System.Collections.Generic.List[int]
    $multiRows = New-Object System.Collections.Generic.List[int]
    $changes = New-Object System.Collections.Generic.List[object]
    $stat = @{ filled = 0; pending = 0; multi = 0; noAnswer = 0; notFound = 0; changes = $changes }
    $now = Get-Date -Format 'yyyy-MM-dd HH:mm'

    for ($i = 0; $i -lt $count; $i++) {
        $existing = $current[($i + 1), 1]
        # 新しい回答が見つからない行は消さずに今の値を残す
        $out[$i, 0] = $existing
        $list = @("$($seibans[($i + 1), 1])" -split '/' | ForEach-Object { Normalize-Key $_ } | Where-Object { $_ })
        if ($list.Count -eq 0) { continue }
        $max = $null; $pending = $false; $found = $false
        foreach ($s in $list) {
            if (-not $answers.ContainsKey($s)) { continue }
            $found = $true
            $e = $answers[$s]
            if ($e.pending) { $pending = $true }
            if ($null -ne $e.max -and ($null -eq $max -or $e.max -gt $max)) { $max = $e.max }
        }
        if (-not $found) { $stat.notFound++; continue }
        if ($null -eq $max) { $stat.noAnswer++; continue }
        $new = Format-Date $max
        $out[$i, 0] = $new
        $stat.filled++
        if ($pending) { $stat.pending++; $pendingRows.Add($first + $i) }
        elseif ($list.Count -gt 1) { $stat.multi++; $multiRows.Add($first + $i) }

        $k = $keys[$i]
        if (-not $k) { continue }
        $prev = "$existing"; $prevTime = ''
        if ($history.ContainsKey($k)) { $prev = $history[$k].value; $prevTime = $history[$k].time }
        $kind = ''
        if (-not $prev) { $kind = '新規' }
        elseif ($prev -ne $new) { $kind = '変更' }
        if ($kind) {
            $days = ''
            $p = [datetime]::MinValue; $n = [datetime]::MinValue
            if ([datetime]::TryParseExact($prev, 'yyyy-MM-dd', $null, 'None', [ref]$p) -and
                [datetime]::TryParseExact($new, 'yyyy-MM-dd', $null, 'None', [ref]$n)) { $days = ($n - $p).Days }
            $changes.Add([pscustomobject]@{ kind = $kind; row = $first + $i; key = $k; seiban = "$($seibans[($i + 1), 1])"; prev = $prev; new = $new; days = $days; prevTime = $prevTime; pending = $pending; multi = ($list.Count -gt 1) })
        }
        $history[$k] = @{ value = $new; seiban = "$($seibans[($i + 1), 1])"; time = $now }
    }

    # 前回の色・文字色・メモを消してから付け直す
    $range.Interior.ColorIndex = -4142
    $range.Font.ColorIndex = -4105
    $range.Font.Bold = $false
    [void]$range.ClearComments()
    $range.NumberFormat = '@'
    $range.Value2 = $out
    foreach ($r in $pendingRows) { $ws.Cells.Item($r, $col).Interior.Color = $PendingColor }
    foreach ($r in $multiRows) { $ws.Cells.Item($r, $col).Interior.Color = $MultiColor }
    # 赤塗り・水色塗りはオレンジ・黄色より優先するので、その内容はメモに書く
    foreach ($c in $changes) {
        $cell = $ws.Cells.Item($c.row, $col)
        $cell.Interior.Color = if ($c.kind -eq '変更') { $ChangedColor } else { $AddedColor }
        $cell.Font.Color = $MarkFontColor
        $cell.Font.Bold = $true
        $text = if ($c.kind -eq '変更') {
            $when = if ($c.prevTime) { "（前回記録 $($c.prevTime)）" } else { '' }
            "納期回答が変わりました`n前回 $($c.prev) → 今回 $($c.new) $(Format-Days $c.days)$when"
        } else {
            "納期回答が新しく入りました`n今回 $($c.new)"
        }
        if ($c.pending) { $text += "`n※回答納期が未入力の発注明細があります" }
        if ($c.multi) { $text += "`n※FD製番が複数あります" }
        [void]$cell.AddComment($text)
    }
    Release $range
    return $stat
}

function Format-Days($days) {
    if ("$days" -eq '') { return '' }
    if ($days -gt 0) { return "+$($days)日遅れ" }
    if ($days -lt 0) { return "$(-$days)日早まり" }
    return ''
}

function Invoke-Fill([string]$zanzanPath, [string]$orderPath, [string[]]$ledgerPaths, [bool]$overwrite, [scriptblock]$log) {
    if (-not (Test-Path -LiteralPath $zanzanPath)) { throw "注残リストが見つかりません: $zanzanPath" }
    if ($orderPath -and -not (Test-Path -LiteralPath $orderPath)) { throw "発注仕入状況一覧表が見つかりません: $orderPath" }

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

        $answers = $null
        if ($orderPath) { $answers = Read-Answers $xl $orderPath $log }
        else { & $log '発注仕入状況一覧表が指定されていないため、納期回答は入れません' }

        # ---- 注残リストに書き込む ----
        $wb = $xl.Workbooks.Open($zanzanPath, 0, $false)
        try {
            if ($wb.ReadOnly) { throw '注残リストが他で開かれています。Excelで閉じてから実行してください。' }
            $ws = $wb.Worksheets.Item(1)
            $h = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ order = '^納入指示番号$'; seiban = '^FD製番$'; date = '^部品納期$'; answer = '^納期回答$' }
            if (-not $h) { throw '注残リストに「納入指示番号」「FD製番」「部品納期」「納期回答」の見出しが見つかりません。' }
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
            $s3 = $null
            $folder = Split-Path $zanzanPath
            $historyPath = Join-Path $folder '納期回答履歴.csv'
            if ($answers) {
                $history = Read-History $historyPath
                $s3 = Fill-Answer $ws $first $count $h.seiban $h.answer $answers $keys $history
            }

            # 上書き前の状態をバックアップしてから保存
            $dir = Join-Path (Split-Path $zanzanPath) 'バックアップ'
            if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
            $backup = Join-Path $dir ("{0}_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($zanzanPath), (Get-Date -Format 'yyyyMMdd_HHmmss'), [IO.Path]::GetExtension($zanzanPath))
            Copy-Item -LiteralPath $zanzanPath -Destination $backup
            $wb.Save()

            # 注残リストを保存できてから、履歴と変更一覧を書き出す
            $changePath = ''
            if ($s3) {
                Save-History $historyPath $history
                if ($s3.changes.Count -gt 0) {
                    $info = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ office = '^事業所名$'; model = '^機種名$'; item = '^品名$' }
                    $changeDir = Join-Path $folder '納期回答変更'
                    if (-not (Test-Path -LiteralPath $changeDir)) { [void](New-Item -ItemType Directory -Path $changeDir) }
                    $changePath = Join-Path $changeDir ("納期回答変更_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
                    $s3.changes | ForEach-Object {
                        [pscustomobject]@{
                            '区分'         = $_.kind
                            '行'           = $_.row
                            '事業所名'     = if ($info) { $ws.Cells.Item($_.row, $info.office).Text } else { '' }
                            '機種名'       = if ($info) { $ws.Cells.Item($_.row, $info.model).Text } else { '' }
                            '品名'         = if ($info) { $ws.Cells.Item($_.row, $info.item).Text } else { '' }
                            '納入指示番号' = $_.key
                            'FD製番'       = $_.seiban
                            '前回納期回答' = $_.prev
                            '今回納期回答' = $_.new
                            '差'           = Format-Days $_.days
                            '前回記録日時' = $_.prevTime
                        }
                    } | Export-Csv -LiteralPath $changePath -NoTypeInformation -Encoding UTF8
                }
            }

            & $log ''
            & $log "対象行              : $count"
            & $log "台帳に無かった行    : $missing"
            & $log "FD製番を入れた行    : $($s1.filled) （うち候補が複数で黄色表示: $($s1.multi)）"
            & $log "部品納期を入れた行  : $($s2.filled) （うち候補が複数で黄色表示: $($s2.multi)）"
            if ($s3) {
                & $log "納期回答を入れた行  : $($s3.filled) （うち未回答の明細ありでオレンジ: $($s3.pending) / 製番が複数で黄色: $($s3.multi)）"
                & $log "  回答納期がまだ無い: $($s3.noAnswer) / 発注仕入状況一覧表に製番が無い: $($s3.notFound)"
                $changed = @($s3.changes | Where-Object { $_.kind -eq '変更' })
                $added = @($s3.changes | Where-Object { $_.kind -eq '新規' })
                & $log "納期回答が前回から変わった行: $($changed.Count) （赤塗り・メモ付き）"
                foreach ($c in ($changed | Select-Object -First 30)) {
                    & $log ("  {0,-12} {1,-18} {2} → {3} {4}" -f $c.key, $c.seiban, $c.prev, $c.new, (Format-Days $c.days))
                }
                if ($changed.Count -gt 30) { & $log "  …ほか $($changed.Count - 30) 行" }
                & $log "納期回答が新しく入った行    : $($added.Count) （水色塗り・メモ付き）"
                if ($changePath) { & $log "  変更・新規の一覧: $changePath" }
            }
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

if (-not $Order) { $Order = Find-OrderFile $Zanzan }

if ($NoGui) {
    Invoke-Fill $Zanzan $Order $Ledgers $Overwrite.IsPresent { param($m) Write-Host $m }
    return
}

# ---------------- 画面 ----------------
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()

$font = New-Object Drawing.Font('Meiryo UI', 9)
$form = New-Object Windows.Forms.Form
$form.Text = '注残リスト 製番転記'
$form.Size = New-Object Drawing.Size(760, 657)
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

$lblO = New-Object Windows.Forms.Label
$lblO.Text = '発注仕入状況一覧表（納期回答の元データ。空欄なら納期回答は入れません）'
$lblO.Location = New-Object Drawing.Point(12, 72); $lblO.AutoSize = $true
$txtO = New-Object Windows.Forms.TextBox
$txtO.Text = $Order
$txtO.Location = New-Object Drawing.Point(12, 92); $txtO.Size = New-Object Drawing.Size(620, 24)
$txtO.Anchor = 'Top,Left,Right'
$btnO = New-Object Windows.Forms.Button
$btnO.Text = '参照...'
$btnO.Location = New-Object Drawing.Point(640, 90); $btnO.Size = New-Object Drawing.Size(90, 27)
$btnO.Anchor = 'Top,Right'
$btnO.Add_Click({
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel|*.xlsx;*.xlsm;*.xls'
    if ($txtZ.Text) { $dlg.InitialDirectory = [IO.Path]::GetDirectoryName($txtZ.Text) }
    if ($dlg.ShowDialog($form) -eq 'OK') { $txtO.Text = $dlg.FileName }
    $dlg.Dispose()
})

$lbl2 = New-Object Windows.Forms.Label
$lbl2.Text = '製番台帳（複数可・ここへファイルをドラッグしても追加できます）'
$lbl2.Location = New-Object Drawing.Point(12, 129); $lbl2.AutoSize = $true
$lst = New-Object Windows.Forms.ListBox
$lst.Location = New-Object Drawing.Point(12, 149); $lst.Size = New-Object Drawing.Size(620, 100)
$lst.Anchor = 'Top,Left,Right'
$lst.HorizontalScrollbar = $true
$lst.SelectionMode = 'MultiExtended'
foreach ($l in $Ledgers) { [void]$lst.Items.Add($l) }

function Add-Ledgers([string[]]$paths) {
    foreach ($f in $paths) {
        if ($f -match '\.(xlsx|xlsm|xls)$' -and -not $lst.Items.Contains($f)) { [void]$lst.Items.Add($f) }
    }
    Save-Settings $txtZ.Text @($lst.Items)
}

# エクスプローラーから台帳ファイルをドラッグして追加できるようにする
$lst.AllowDrop = $true
$lst.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = 'Copy' }
})
$lst.Add_DragDrop({ Add-Ledgers $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop) })

$btnAdd = New-Object Windows.Forms.Button
$btnAdd.Text = '追加...'
$btnAdd.Location = New-Object Drawing.Point(640, 149); $btnAdd.Size = New-Object Drawing.Size(90, 27)
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
$btnDel.Location = New-Object Drawing.Point(640, 182); $btnDel.Size = New-Object Drawing.Size(90, 27)
$btnDel.Anchor = 'Top,Right'
$btnDel.Add_Click({
    foreach ($item in @($lst.SelectedItems)) { $lst.Items.Remove($item) }
    Save-Settings $txtZ.Text @($lst.Items)
})

$chk = New-Object Windows.Forms.CheckBox
$chk.Text = 'すでにFD製番・部品納期が入っている行も上書きする（納期回答は毎回最新に入れ直します）'
$chk.Checked = $Overwrite.IsPresent
$chk.Location = New-Object Drawing.Point(12, 259); $chk.AutoSize = $true

$btnRun = New-Object Windows.Forms.Button
$btnRun.Text = '実行'
$btnRun.Location = New-Object Drawing.Point(640, 255); $btnRun.Size = New-Object Drawing.Size(90, 32)
$btnRun.Anchor = 'Top,Right'

$txtLog = New-Object Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Location = New-Object Drawing.Point(12, 297); $txtLog.Size = New-Object Drawing.Size(718, 305)
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
        Save-Settings $txtZ.Text @($lst.Items)
        Invoke-Fill $txtZ.Text $txtO.Text.Trim() @($lst.Items) $chk.Checked $log
        [Windows.Forms.MessageBox]::Show($form, '完了しました。', '製番転記') | Out-Null
    } catch {
        & $log "エラー: $($_.Exception.Message)"
        [Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'エラー', 'OK', 'Error') | Out-Null
    } finally {
        $form.Cursor = 'Default'
        $btnRun.Enabled = $true
    }
})

$form.Controls.AddRange(@($lbl1, $txtZ, $btnZ, $lblO, $txtO, $btnO, $lbl2, $lst, $btnAdd, $btnDel, $chk, $btnRun, $txtLog))
$form.Add_FormClosing({ try { Save-Settings $txtZ.Text @($lst.Items) } catch { } })
[void]$form.ShowDialog()

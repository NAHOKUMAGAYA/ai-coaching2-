# 注残リストの「納入指示番号」から製番台帳を引いて「FD製番」「部品納期」列に書き込み、
# その製番の発注明細(発注仕入状況一覧表)から一番遅い回答納期を「納期回答」列に、回答備考を「コメント」列に書き込む
# 使い方: 製番転記.bat をダブルクリック
#   コマンドで実行する場合: powershell -File 製番転記.ps1 -NoGui [-Zanzan パス] [-Order パス] [-Overwrite]
param(
    [string]$Zanzan = 'C:\Users\n-kumagaya\Desktop\ai-coaching2-\製番転記_納期回答\データ\注残リスト.xlsx',
    [string]$Order = '',   # 発注仕入状況一覧表。省略時は注残リストと同じフォルダの一番新しいもの
    [string]$Exclude = 'C:\Users\n-kumagaya\Desktop\ai-coaching2-\製番転記_納期回答\データ\指示区分変更依頼追記分(20260911)土岐含む.xlsx',   # 対象外品番リスト。空ならこの条件は使わない
    [string[]]$Ledgers = @(
        '\\192.168.11.200\Share\12.製番台帳\43期製番台帳\（新）4374(AMD単品(富士宮)).xlsx',
        '\\192.168.11.200\Share\12.製番台帳\43期製番台帳\（新）4366(AMD単品(土岐)).xlsx'
    ),
    [switch]$Overwrite,
    [switch]$NoGui
)

$ErrorActionPreference = 'Stop'
$SkipSheets = @('記入方法', 'リスト', '休日')
# 納期回答・回答備考を見るときに無視する発注先番号（この発注先の明細は回答納期が入っていても使わない）
$IgnoreSuppliers = @('FSH', 'FZH', 'FZH2', 'NK', 'SFD', 'SFH', 'SFZ', 'SM', 'SMF', 'SPH', 'SPM', 'SPP', 'STK')
# FD製番に「対象外」と入れる条件（どれか1つに当てはまる行）:
#   ユーザ欄1が J 以外 / 納入指示区分が JIT（手配JIT は含まない）/ 品番が対象外品番リストのこのシートにある
$ExcludeText = '対象外'
$ExcludeSheet = '納期回答対象外'
# 上の条件に当てはまっても対象外にしない品番
$KeepParts = @('EK0393')

# 画面で選んだ注残リストと製番台帳を覚えておき、次回もそのまま使う
$SettingsPath = Join-Path $PSScriptRoot '製番転記_設定.json'
if (Test-Path -LiteralPath $SettingsPath) {
    $saved = Get-Content -LiteralPath $SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($saved.Zanzan -and -not $PSBoundParameters.ContainsKey('Zanzan')) { $Zanzan = $saved.Zanzan }
    if ($null -ne $saved.Exclude -and -not $PSBoundParameters.ContainsKey('Exclude')) { $Exclude = $saved.Exclude }
    if ($saved.Ledgers -and -not $PSBoundParameters.ContainsKey('Ledgers')) { $Ledgers = @($saved.Ledgers) }
}

function Save-Settings([string]$zanzanPath, [string]$excludePath, [string[]]$ledgerPaths) {
    ConvertTo-Json -InputObject ([ordered]@{ Zanzan = $zanzanPath; Exclude = $excludePath; Ledgers = @($ledgerPaths) }) |
        Set-Content -LiteralPath $SettingsPath -Encoding UTF8
}
# セルの色分け（Excelの色は BGR の数値）
#   製番が2つ以上 : 黄色塗り・赤太字（ほかより優先）
#   納期回答の新規: 水色塗り・白字 / 変更: 赤塗り・白字 / 前回と同じ: 塗りなし・黒字
$MultiColor = 65535       # 黄色
$MultiFontColor = 255     # 赤
$ChangedColor = 255       # 赤
$AddedColor = 15773696    # 水色
$MarkFontColor = 16777215 # 白

function Set-MultiStyle($cell) {
    $cell.Interior.Color = $MultiColor
    $cell.Font.Color = $MultiFontColor
    $cell.Font.Bold = $true
}

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

# 注残リストの1列を、納入指示番号をキーに対応表から埋める。候補が複数ある行は「/」でつないで黄色塗り・赤太字にする。
# 対象外の行($excluded)は台帳を引かず、$excludedText があればそれを入れる（無ければ空のまま）
function Fill-Column($ws, [int]$first, [int]$col, [string[]]$keys, [hashtable]$lookup, [bool]$overwrite, [bool[]]$excluded, [string]$excludedText) {
    $count = $keys.Count
    $range = $ws.Range($ws.Cells.Item($first, $col), $ws.Cells.Item($first + $count - 1, $col))
    $current = $range.Value2
    $out = New-Object 'object[,]' $count, 1
    $multiRows = New-Object System.Collections.Generic.List[int]
    $singleRows = New-Object System.Collections.Generic.List[int]
    $stat = @{ filled = 0; multi = 0; kept = 0; excluded = 0 }

    for ($i = 0; $i -lt $count; $i++) {
        $existing = $current[($i + 1), 1]
        $out[$i, 0] = $existing
        if (-not $overwrite -and "$existing" -ne '') {
            # 残した行も、前の色分けやダウンロード時の赤太字のままにならないよう色分けし直す
            $stat.kept++
            if ("$existing" -like '*/*') { $multiRows.Add($first + $i) } else { $singleRows.Add($first + $i) }
            continue
        }
        if ($excluded[$i]) {
            if ($excludedText) { $out[$i, 0] = $excludedText; $stat.excluded++; $singleRows.Add($first + $i) }
            continue
        }
        $k = $keys[$i]
        if (-not $k -or -not $lookup.ContainsKey($k)) { continue }
        $out[$i, 0] = ($lookup[$k] -join '/')
        $stat.filled++
        if ($lookup[$k].Count -gt 1) { $stat.multi++; $multiRows.Add($first + $i) }
        else { $singleRows.Add($first + $i) }
    }

    # 製番や日付が数値に変換されないよう文字列書式にしてから一括書き込み
    $range.NumberFormat = '@'
    $range.Value2 = $out
    foreach ($r in $multiRows) { Set-MultiStyle $ws.Cells.Item($r, $col) }
    # 製番が1つの行は塗りなし・黒字にする
    foreach ($r in $singleRows) {
        $cell = $ws.Cells.Item($r, $col)
        $cell.Interior.ColorIndex = -4142
        $cell.Font.ColorIndex = -4105
        $cell.Font.Bold = $false
    }
    Release $range
    return $stat
}

# 対象外品番リストの「納期回答対象外」シートから品番を読む
function Read-ExcludeItems($xl, [string]$path, [scriptblock]$log) {
    & $log "対象外品番リスト読込: $(Split-Path $path -Leaf)"
    $items = New-Object 'System.Collections.Generic.HashSet[string]'
    $wb = $xl.Workbooks.Open($path, 0, $true)
    try {
        $ws = $null
        foreach ($s in $wb.Worksheets) { if ($s.Name -eq $ExcludeSheet) { $ws = $s } }
        if (-not $ws) { throw "対象外品番リストに「$ExcludeSheet」シートが見つかりません。" }
        $h = Find-Header $ws.Range('A1:Z5').Value2 5 @{ item = '^品番$' }
        if (-not $h) { throw "対象外品番リストの「$ExcludeSheet」シートに「品番」の見出しが見つかりません。" }
        $last = $ws.Cells.Item($ws.Rows.Count, $h.item).End(-4162).Row
        if ($last -gt $h.row) {
            $v = $ws.Range($ws.Cells.Item(1, $h.item), $ws.Cells.Item($last, $h.item)).Value2
            for ($r = $h.row + 1; $r -le $last; $r++) {
                $k = Normalize-Key "$($v[$r, 1])"
                if ($k) { [void]$items.Add($k) }
            }
        }
        & $log "  シート[$ExcludeSheet] 品番 $($items.Count) 件"
    } finally {
        $wb.Close($false)
        Release $wb
    }
    # 先頭の「,」は、HashSet がばらされて配列にならないようにするため
    return , $items
}

# FD製番に「対象外」と入れる行を決める。ユーザ欄1が J 以外・納入指示区分が JIT・品番が対象外品番リストにある、のどれかに当てはまる行。
# 見出しが無い条件は使わない
function Find-Excluded($ws, [int]$first, [int]$count, $items, [scriptblock]$log) {
    $last = $first + $count - 1
    $col = {
        param($pattern, $label)
        $h = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ c = $pattern }
        if (-not $h) { & $log "※注残リストに「$label」の見出しが無いため、この条件では対象外にしません"; return }
        , $ws.Range($ws.Cells.Item($first, $h.c), $ws.Cells.Item($last, $h.c)).Value2
    }
    $users = & $col '^ユーザ欄[1１]$' 'ユーザ欄1'
    $kinds = & $col '^納入指示区分$' '納入指示区分'
    $parts = & $col '^品番$' '品番'

    $stat = @{ flags = (New-Object 'bool[]' $count); rows = 0; user = 0; jit = 0; item = 0 }
    for ($i = 1; $i -le $count; $i++) {
        $hit = $false
        $part = if ($null -ne $parts) { Normalize-Key "$($parts[$i, 1])" } else { '' }
        if ($part -and $KeepParts -contains $part) { continue }
        if ($null -ne $users -and (Normalize-Key "$($users[$i, 1])") -ne 'J') { $stat.user++; $hit = $true }
        # 「1:ＪＩＴ」のように番号が付いているので外して比べる
        if ($null -ne $kinds -and ((Normalize-Key "$($kinds[$i, 1])") -replace '^\d+:', '') -eq 'JIT') { $stat.jit++; $hit = $true }
        if ($null -ne $items -and $part -and $items.Contains($part)) { $stat.item++; $hit = $true }
        if ($hit) { $stat.flags[$i - 1] = $true; $stat.rows++ }
    }
    return $stat
}

# 注残リストに書き込む列（見出し → 探すときのパターン）。無い列は納入指示番号の右にこの順で挿入する
$FillColumns = [ordered]@{ 'FD製番' = '^FD製番$'; '部品納期' = '^部品納期$'; '納期回答' = '^納期回答$'; 'コメント' = '^(コメント|回答備考)$' }

function Add-MissingColumns($ws, [scriptblock]$log) {
    $h = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ order = '^納入指示番号$' }
    if (-not $h) { return }   # 見出しが無いエラーはこのあとで出す
    $col = $h.order
    $added = New-Object System.Collections.Generic.List[string]
    foreach ($name in $FillColumns.Keys) {
        $found = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ c = $FillColumns[$name] }
        if ($found) { $col = $found.c; continue }
        $col++
        # 見出しの書式は左隣(納入指示番号など)と同じにする
        [void]$ws.Columns.Item($col).Insert(-4161, 0)
        # Value2 に文字を1つ入れると、あとの列への一括書き込み(Value2 = 配列)が型エラーになるため Value で入れる
        $ws.Cells.Item($h.row, $col).Value = $name
        $ws.Columns.Item($col).ColumnWidth = 12
        $added.Add($name)
    }
    if ($added.Count) { & $log "注残リストに列を挿入しました: $($added -join '・')" }
}

function Find-OrderFile([string]$zanzanPath) {
    $dir = Split-Path $zanzanPath
    if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { return '' }
    $f = Get-ChildItem -LiteralPath $dir -Filter '発注仕入状況一覧表*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($f) { return $f.FullName }
    return ''
}

# 発注仕入状況一覧表の明細のうち「発注明細」シートに出す列（見出し → 発注明細シートでの見出し）
$DetailColumns = [ordered]@{
    '発注番号' = '発注番号'; '発注先番号' = '発注先番号'; '発注先名称' = '発注先名称'; '品名規格' = '品名規格'; '型式材質' = '型式材質'
    '発注数計' = '発注数'; '仕入数計' = '仕入数'; '仕入残' = '仕入残'; '仕入状況' = '仕入状況'
    '発注日' = '発注日'; '部品納期' = '部品納期'; '回答納期' = '回答納期'
}

# 発注仕入状況一覧表から 製造番号 → 一番遅い回答納期・未回答の明細の有無 を集計する。
# $detail には「発注明細」シート用に、製造番号ごとの明細の行番号と各列の値を入れる
function Read-Answers($xl, [string]$path, [scriptblock]$log, [hashtable]$detail) {
    & $log "発注仕入状況一覧表読込: $(Split-Path $path -Leaf)"
    $answers = @{}
    $wb = $xl.Workbooks.Open($path, 0, $true)
    try {
        $ws = $wb.Worksheets.Item(1)
        $h = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ supplier = '^発注先番号$'; seiban = '^製造番号$'; answer = '^回答納期$'; remark = '^回答備考$'; remain = '^仕入残$' }
        if (-not $h) { throw '発注仕入状況一覧表に「発注先番号」「製造番号」「回答納期」「回答備考」「仕入残」の見出しが見つかりません。' }
        $last = $ws.Cells.Item($ws.Rows.Count, $h.seiban).End(-4162).Row
        $sp = $ws.Range($ws.Cells.Item(1, $h.supplier), $ws.Cells.Item($last, $h.supplier)).Value2
        $sb = $ws.Range($ws.Cells.Item(1, $h.seiban), $ws.Cells.Item($last, $h.seiban)).Value2
        $an = $ws.Range($ws.Cells.Item(1, $h.answer), $ws.Cells.Item($last, $h.answer)).Value2
        $bk = $ws.Range($ws.Cells.Item(1, $h.remark), $ws.Cells.Item($last, $h.remark)).Value2
        $rm = $ws.Range($ws.Cells.Item(1, $h.remain), $ws.Cells.Item($last, $h.remain)).Value2

        $headRow = $ws.Range($ws.Cells.Item($h.row, 1), $ws.Cells.Item($h.row, 80)).Value2
        $detail.cols = @{}
        foreach ($name in $DetailColumns.Keys) {
            for ($c = 1; $c -le 80; $c++) {
                if (("$($headRow[1, $c])" -replace '\s', '') -eq $name) {
                    $detail.cols[$name] = $ws.Range($ws.Cells.Item(1, $c), $ws.Cells.Item($last, $c)).Value2
                    break
                }
            }
        }
        $detail.rows = @{}
        $detail.remarks = New-Object 'string[]' ($last + 1)

        $prevRemark = ''
        $ignored = 0
        for ($r = $h.row + 1; $r -le $last; $r++) {
            # 「〃」はすぐ上の行(同じ発注先)の回答備考と同じという意味なので置き換える
            $remark = "$($bk[$r, 1])".Trim()
            if ($remark -match '^[〃″"]+$' -and $r -gt $h.row + 1 -and "$($sp[$r, 1])" -eq "$($sp[($r - 1), 1])") { $remark = $prevRemark }
            $prevRemark = $remark
            $detail.remarks[$r] = $remark

            $k = Normalize-Key "$($sb[$r, 1])"
            if (-not $k) { continue }
            if (-not $detail.rows.ContainsKey($k)) { $detail.rows[$k] = New-Object System.Collections.Generic.List[int] }
            $detail.rows[$k].Add($r)
            if ($IgnoreSuppliers -contains (Normalize-Key "$($sp[$r, 1])")) { $ignored++; continue }
            if (-not $answers.ContainsKey($k)) {
                $answers[$k] = @{
                    max = $null; pending = $false
                    maxRemarks = New-Object System.Collections.Generic.List[string]      # 一番遅い回答納期の明細の回答備考
                    pendingRemarks = New-Object System.Collections.Generic.List[string]  # 回答納期が空で入荷が残る明細の回答備考
                }
            }
            $e = $answers[$k]
            $a = $an[$r, 1]
            if ($a -is [double]) {
                if ($null -eq $e.max -or $a -gt $e.max) { $e.max = $a; $e.maxRemarks.Clear() }
                if ($a -eq $e.max -and $remark -and -not $e.maxRemarks.Contains($remark)) { $e.maxRemarks.Add($remark) }
            } elseif ($rm[$r, 1] -is [double] -and $rm[$r, 1] -gt 0) {
                # 入荷が残っているのに回答納期が空の明細
                $e.pending = $true
                if ($remark -and -not $e.pendingRemarks.Contains($remark)) { $e.pendingRemarks.Add($remark) }
            }
        }
        & $log "  発注明細 $($last - $h.row) 行（うち無視する発注先 $ignored 行） / 製造番号 $($answers.Count) 件"
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
            $history[$row.'納入指示番号'] = @{ value = $row.'納期回答'; seiban = $row.'FD製番'; time = $row.'記録日時'; remark = "$($row.'回答備考')" }
        }
    }
    return $history
}

function Save-History([string]$path, [hashtable]$history) {
    $history.GetEnumerator() | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{ '納入指示番号' = $_.Name; 'FD製番' = $_.Value.seiban; '納期回答' = $_.Value.value; '記録日時' = $_.Value.time; '回答備考' = $_.Value.remark }
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
    $stat = @{ filled = 0; pending = 0; multi = 0; noAnswer = 0; notFound = 0; cleared = 0; changes = $changes }
    $now = Get-Date -Format 'yyyy-MM-dd HH:mm'

    for ($i = 0; $i -lt $count; $i++) {
        $existing = $current[($i + 1), 1]
        $out[$i, 0] = $existing
        $list = @("$($seibans[($i + 1), 1])" -split '/' | ForEach-Object { Normalize-Key $_ } | Where-Object { $_ -and $_ -ne $ExcludeText })
        if ($list.Count -eq 0) { continue }
        $max = $null; $pending = $false; $found = $false
        foreach ($s in $list) {
            if (-not $answers.ContainsKey($s)) { continue }
            $found = $true
            $e = $answers[$s]
            if ($e.pending) { $pending = $true }
            if ($null -ne $e.max -and ($null -eq $max -or $e.max -gt $max)) { $max = $e.max }
        }
        $k = $keys[$i]
        if (-not $found -or $null -eq $max) {
            if (-not $found) { $stat.notFound++ } else { $stat.noAnswer++ }
            # 回答が無くなった行は、前回このツールが入れた値のままなら消す（無視する発注先の回答などが残らないように）。
            # 自分で入力した値(履歴と違う値)は残す
            if ($k -and "$existing" -ne '' -and $history.ContainsKey($k) -and $history[$k].value -eq "$existing") {
                $out[$i, 0] = $null
                $history[$k].value = ''
                $history[$k].time = $now
                $stat.cleared++
            }
            continue
        }
        $new = Format-Date $max
        $out[$i, 0] = $new
        $stat.filled++
        if ($pending) { $stat.pending++; $pendingRows.Add($first + $i) }
        if ($list.Count -gt 1) { $stat.multi++; $multiRows.Add($first + $i) }

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
        if (-not $history.ContainsKey($k)) { $history[$k] = @{ remark = '' } }
        $history[$k].value = $new
        $history[$k].seiban = "$($seibans[($i + 1), 1])"
        $history[$k].time = $now
    }

    # 前回の色・文字色・メモを消してから付け直す
    $range.Interior.ColorIndex = -4142
    $range.Font.ColorIndex = -4105
    $range.Font.Bold = $false
    [void]$range.ClearComments()
    $range.NumberFormat = '@'
    $range.Value2 = $out
    # 新規は水色塗り・白字、変更は赤塗り・白字。製番が複数の行は黄色塗り・赤太字を優先し、新規・変更はメモで分かるようにする
    $notes = @{}
    foreach ($c in $changes) {
        if (-not $c.multi) {
            $cell = $ws.Cells.Item($c.row, $col)
            $cell.Interior.Color = if ($c.kind -eq '変更') { $ChangedColor } else { $AddedColor }
            $cell.Font.Color = $MarkFontColor
        }
        $notes[$c.row] = if ($c.kind -eq '変更') {
            $when = if ($c.prevTime) { "（前回記録 $($c.prevTime)）" } else { '' }
            "納期回答が変わりました`n前回 $($c.prev) → 今回 $($c.new) $(Format-Days $c.days)$when"
        } else {
            "納期回答が新しく入りました`n今回 $($c.new)"
        }
    }
    foreach ($r in $multiRows) {
        Set-MultiStyle $ws.Cells.Item($r, $col)
        $notes[$r] = (@($notes[$r], '※FD製番が複数あります') | Where-Object { $_ }) -join "`n"
    }
    foreach ($r in $pendingRows) {
        $notes[$r] = (@($notes[$r], '※回答納期が未入力の発注明細があります') | Where-Object { $_ }) -join "`n"
    }
    foreach ($r in $notes.Keys) { [void]$ws.Cells.Item($r, $col).AddComment($notes[$r]) }
    Release $range
    return $stat
}

# FD製番ごとの回答備考を「コメント」列に毎回入れ直す。
# 一番遅い回答納期の明細の備考に、回答納期が空で入荷が残る明細の備考を続ける。
# 備考が無くなった行は、前回このツールが入れた値のままなら消し、自分で入力した値は残す
function Fill-Remark($ws, [int]$first, [int]$count, [int]$seibanCol, [int]$col, [hashtable]$answers, [string[]]$keys, [hashtable]$history) {
    $seibans = $ws.Range($ws.Cells.Item($first, $seibanCol), $ws.Cells.Item($first + $count - 1, $seibanCol)).Value2
    $range = $ws.Range($ws.Cells.Item($first, $col), $ws.Cells.Item($first + $count - 1, $col))
    $out = $range.Value2
    $stat = @{ filled = 0; cleared = 0 }

    for ($i = 1; $i -le $count; $i++) {
        $k = $keys[($i - 1)]
        $list = @("$($seibans[$i, 1])" -split '/' | ForEach-Object { Normalize-Key $_ } | Where-Object { $_ -and $answers.ContainsKey($_) })
        $max = $null
        foreach ($s in $list) { $m = $answers[$s].max; if ($null -ne $m -and ($null -eq $max -or $m -gt $max)) { $max = $m } }
        $remarks = New-Object System.Collections.Generic.List[string]
        foreach ($s in $list) {
            if ($null -ne $max -and $answers[$s].max -eq $max) { foreach ($t in $answers[$s].maxRemarks) { if (-not $remarks.Contains($t)) { $remarks.Add($t) } } }
        }
        foreach ($s in $list) { foreach ($t in $answers[$s].pendingRemarks) { if (-not $remarks.Contains($t)) { $remarks.Add($t) } } }

        if ($remarks.Count -eq 0) {
            $existing = "$($out[$i, 1])"
            if ($k -and $existing -ne '' -and $history.ContainsKey($k) -and $history[$k].remark -eq $existing) {
                $out[$i, 1] = $null
                $history[$k].remark = ''
                $stat.cleared++
            }
            continue
        }
        $text = $remarks -join ' / '
        $out[$i, 1] = $text
        $stat.filled++
        if ($k) {
            if (-not $history.ContainsKey($k)) { $history[$k] = @{ value = ''; seiban = "$($seibans[$i, 1])"; time = (Get-Date -Format 'yyyy-MM-dd HH:mm') } }
            $history[$k].remark = $text
        }
    }

    # 列を挿入したときに左隣(納期回答)からコピーされた色などが残らないよう、書式を戻してから書き込む。
    # 「7/15」などが日付に変換されないよう文字列書式にする
    $range.Interior.ColorIndex = -4142
    $range.Font.ColorIndex = -4105
    $range.Font.Bold = $false
    $range.NumberFormat = '@'
    $range.Value2 = $out
    Release $range
    return $stat
}

# 注残リストの行ごとに、FD製番の発注明細を1行ずつ「発注明細」シートに並べる（毎回作り直す）。
# 納期回答の元になった明細には★、無視する発注先の明細には「対象外」を付ける
function Write-DetailSheet($wb, $ws, [int]$first, [int]$count, [int]$seibanCol, [int]$orderCol, [int]$answerCol, [hashtable]$answers, [hashtable]$detail) {
    $sheetName = '発注明細'
    $info = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ office = '^事業所名$'; item = '^品名$' }
    $last = $first + $count - 1
    # 先頭の「,」は、2次元配列が平たくならないようにするため
    $col = { param($c) , $ws.Range($ws.Cells.Item($first, $c), $ws.Cells.Item($last, $c)).Value2 }
    $seibans = & $col $seibanCol
    $orders = & $col $orderCol
    $answerVals = & $col $answerCol
    $offices = $null; $items = $null
    if ($info) { $offices = & $col $info.office; $items = & $col $info.item }

    # 出す行（注残の行, 発注明細の行, その注残の行の一番遅い回答納期）を先に集める
    $pairs = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $count; $i++) {
        $list = @("$($seibans[$i, 1])" -split '/' | ForEach-Object { Normalize-Key $_ } | Where-Object { $_ })
        $max = $null
        foreach ($s in $list) {
            if ($answers.ContainsKey($s)) { $m = $answers[$s].max; if ($null -ne $m -and ($null -eq $max -or $m -gt $max)) { $max = $m } }
        }
        foreach ($s in $list) {
            if (-not $detail.rows.ContainsKey($s)) { continue }
            foreach ($r in $detail.rows[$s]) { $pairs.Add(@($i, $r, $max)) }
        }
    }

    $heads = @('注残の行', '事業所名', '納入指示番号', '品名', 'FD製番', '納期回答') + @($DetailColumns.Values) + @('回答備考', '区分')
    $numberHeads = @('注残の行', '発注数', '仕入数', '仕入残')
    $dateHeads = @('発注日', '部品納期', '回答納期')
    $out = New-Object 'object[,]' ($pairs.Count + 1), $heads.Count
    for ($c = 0; $c -lt $heads.Count; $c++) { $out[0, $c] = $heads[$c] }
    $n = 1
    foreach ($p in $pairs) {
        $i = $p[0]; $r = $p[1]; $max = $p[2]
        $out[$n, 0] = $first + $i - 1
        $out[$n, 1] = if ($offices) { "$($offices[$i, 1])" } else { '' }
        $out[$n, 2] = "$($orders[$i, 1])"
        $out[$n, 3] = if ($items) { "$($items[$i, 1])" } else { '' }
        $out[$n, 4] = "$($seibans[$i, 1])"
        $out[$n, 5] = "$($answerVals[$i, 1])"
        $c = 6
        foreach ($name in $DetailColumns.Keys) {
            $v = if ($detail.cols.ContainsKey($name)) { $detail.cols[$name][$r, 1] } else { $null }
            $out[$n, $c] = if ($dateHeads -contains $DetailColumns[$name]) { Format-Date $v }
                           elseif ($numberHeads -contains $DetailColumns[$name]) { $v }
                           else { "$v".Trim() }
            $c++
        }
        $out[$n, $c] = $detail.remarks[$r]
        $supplier = if ($detail.cols.ContainsKey('発注先番号')) { Normalize-Key "$($detail.cols['発注先番号'][$r, 1])" } else { '' }
        $answer = if ($detail.cols.ContainsKey('回答納期')) { $detail.cols['回答納期'][$r, 1] } else { $null }
        $out[$n, ($c + 1)] = if ($IgnoreSuppliers -contains $supplier) { '対象外の発注先' }
                             elseif ($null -ne $max -and $answer -is [double] -and $answer -eq $max) { '★納期回答の元' }
                             else { '' }
        $n++
    }

    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $sheetName) { $s.Delete() } }
    $ds = $wb.Worksheets.Add([Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $ds.Name = $sheetName
    $rows = $pairs.Count + 1
    for ($c = 1; $c -le $heads.Count; $c++) {
        if ($numberHeads -notcontains $heads[$c - 1]) { $ds.Range($ds.Cells.Item(1, $c), $ds.Cells.Item($rows, $c)).NumberFormat = '@' }
    }
    $all = $ds.Range($ds.Cells.Item(1, 1), $ds.Cells.Item($rows, $heads.Count))
    $all.Value2 = $out
    $head = $ds.Range($ds.Cells.Item(1, 1), $ds.Cells.Item(1, $heads.Count))
    $head.Font.Bold = $true
    $head.Interior.Color = 14277081   # 薄い灰色
    [void]$all.AutoFilter()
    [void]$all.Columns.AutoFit()
    # 元のシートを開いた状態で保存する
    [void]$ws.Activate()
    Release $all; Release $head; Release $ds
    return $pairs.Count
}

function Format-Days($days) {
    if ("$days" -eq '') { return '' }
    if ($days -gt 0) { return "+$($days)日遅れ" }
    if ($days -lt 0) { return "$(-$days)日早まり" }
    return ''
}

function Invoke-Fill([string]$zanzanPath, [string]$orderPath, [string]$excludePath, [string[]]$ledgerPaths, [bool]$overwrite, [scriptblock]$log) {
    if (-not (Test-Path -LiteralPath $zanzanPath)) { throw "注残リストが見つかりません: $zanzanPath" }
    if ($orderPath -and -not (Test-Path -LiteralPath $orderPath)) { throw "発注仕入状況一覧表が見つかりません: $orderPath" }
    if ($excludePath -and -not (Test-Path -LiteralPath $excludePath)) { throw "対象外品番リストが見つかりません: $excludePath" }

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
        $detail = @{}
        if ($orderPath) { $answers = Read-Answers $xl $orderPath $log $detail }
        else { & $log '発注仕入状況一覧表が指定されていないため、納期回答は入れません' }

        $excludeItems = $null
        if ($excludePath) { $excludeItems = Read-ExcludeItems $xl $excludePath $log }
        else { & $log '対象外品番リストが指定されていないため、品番では対象外にしません' }

        # ---- 注残リストに書き込む ----
        $wb = $xl.Workbooks.Open($zanzanPath, 0, $false)
        try {
            if ($wb.ReadOnly) { throw '注残リストが他で開かれています。Excelで閉じてから実行してください。' }
            $ws = $wb.Worksheets.Item(1)
            Add-MissingColumns $ws $log
            $h = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ order = '^納入指示番号$'; seiban = '^FD製番$'; date = '^部品納期$'; answer = '^納期回答$' }
            if (-not $h) { throw '注残リストに「納入指示番号」の見出しが見つかりません。' }
            $first = $h.row + 1
            $last = $ws.Cells.Item($ws.Rows.Count, $h.order).End(-4162).Row
            if ($last -lt $first) { throw '注残リストにデータがありません。' }

            $orders = $ws.Range($ws.Cells.Item($first, $h.order), $ws.Cells.Item($last, $h.order)).Value2
            $count = $last - $first + 1
            $ex = Find-Excluded $ws $first $count $excludeItems $log
            $keys = New-Object 'string[]' $count
            $missing = 0
            for ($i = 0; $i -lt $count; $i++) {
                $keys[$i] = Normalize-Key "$($orders[($i + 1), 1])"
                if ($keys[$i] -and -not $ex.flags[$i] -and -not $map.ContainsKey($keys[$i])) { $missing++ }
            }

            # 対象外の行は FD製番に「対象外」と入れ、部品納期は入れない
            $s1 = Fill-Column $ws $first $h.seiban $keys $map $overwrite $ex.flags $ExcludeText
            $s2 = Fill-Column $ws $first $h.date $keys $dateMap $overwrite $ex.flags ''
            $s3 = $null
            $folder = Split-Path $zanzanPath
            $historyPath = Join-Path $folder '納期回答履歴.csv'
            if ($answers) {
                $history = Read-History $historyPath
                $s3 = Fill-Answer $ws $first $count $h.seiban $h.answer $answers $keys $history
                $rc = Find-Header $ws.Range('A1:BZ5').Value2 5 @{ remark = '^(コメント|回答備考)$' }
                $s4 = $null
                if ($rc) { $s4 = Fill-Remark $ws $first $count $h.seiban $rc.remark $answers $keys $history }
                $detailRows = Write-DetailSheet $wb $ws $first $count $h.seiban $h.order $h.answer $answers $detail
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
            & $log "対象外の条件に当たる行: $($ex.rows) （ユーザ欄1がJ以外 $($ex.user) / 納入指示区分がJIT $($ex.jit) / 対象外品番 $($ex.item)。重複あり）"
            & $log "台帳に無かった行    : $missing （対象外の行は除く）"
            & $log "FD製番を入れた行    : $($s1.filled) （うち候補が複数で黄色: $($s1.multi)）"
            & $log "FD製番に「対象外」を入れた行: $($s1.excluded)"
            & $log "部品納期を入れた行  : $($s2.filled) （うち候補が複数で黄色: $($s2.multi)）"
            if ($s3) {
                & $log "納期回答を入れた行  : $($s3.filled) （うち製番が複数で黄色: $($s3.multi) / 未回答の明細ありをメモ: $($s3.pending)）"
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
                if ($s3.cleared) { & $log "回答が無くなり納期回答を消した行: $($s3.cleared)" }
                if ($s4) { & $log "回答備考を入れた行  : $($s4.filled)$(if ($s4.cleared) { " （備考が無くなり消した行: $($s4.cleared)）" })" }
                else { & $log '注残リストに「コメント」列が無いため、回答備考は入れていません' }
                & $log "「発注明細」シート  : $detailRows 行"
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
    Invoke-Fill $Zanzan $Order $Exclude $Ledgers $Overwrite.IsPresent { param($m) Write-Host $m }
    return
}

# ---------------- 画面 ----------------
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()

$font = New-Object Drawing.Font('Meiryo UI', 9)
$form = New-Object Windows.Forms.Form
$form.Text = '注残リスト 製番転記'
$form.Size = New-Object Drawing.Size(760, 714)
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

$lblE = New-Object Windows.Forms.Label
$lblE.Text = "対象外品番リスト（「$ExcludeSheet」シートの品番はFD製番を「対象外」にします。空欄ならこの条件は使いません）"
$lblE.Location = New-Object Drawing.Point(12, 129); $lblE.AutoSize = $true
$txtE = New-Object Windows.Forms.TextBox
$txtE.Text = $Exclude
$txtE.Location = New-Object Drawing.Point(12, 149); $txtE.Size = New-Object Drawing.Size(620, 24)
$txtE.Anchor = 'Top,Left,Right'
$btnE = New-Object Windows.Forms.Button
$btnE.Text = '参照...'
$btnE.Location = New-Object Drawing.Point(640, 147); $btnE.Size = New-Object Drawing.Size(90, 27)
$btnE.Anchor = 'Top,Right'
$btnE.Add_Click({
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel|*.xlsx;*.xlsm;*.xls'
    if ($txtE.Text) { $dlg.InitialDirectory = [IO.Path]::GetDirectoryName($txtE.Text) }
    if ($dlg.ShowDialog($form) -eq 'OK') { $txtE.Text = $dlg.FileName }
    $dlg.Dispose()
})

$lbl2 = New-Object Windows.Forms.Label
$lbl2.Text = '製番台帳（複数可・ここへファイルをドラッグしても追加できます）'
$lbl2.Location = New-Object Drawing.Point(12, 186); $lbl2.AutoSize = $true
$lst = New-Object Windows.Forms.ListBox
$lst.Location = New-Object Drawing.Point(12, 206); $lst.Size = New-Object Drawing.Size(620, 100)
$lst.Anchor = 'Top,Left,Right'
$lst.HorizontalScrollbar = $true
$lst.SelectionMode = 'MultiExtended'
foreach ($l in $Ledgers) { [void]$lst.Items.Add($l) }

function Add-Ledgers([string[]]$paths) {
    foreach ($f in $paths) {
        if ($f -match '\.(xlsx|xlsm|xls)$' -and -not $lst.Items.Contains($f)) { [void]$lst.Items.Add($f) }
    }
    Save-Settings $txtZ.Text $txtE.Text.Trim() @($lst.Items)
}

# エクスプローラーから台帳ファイルをドラッグして追加できるようにする
$lst.AllowDrop = $true
$lst.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = 'Copy' }
})
$lst.Add_DragDrop({ Add-Ledgers $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop) })

$btnAdd = New-Object Windows.Forms.Button
$btnAdd.Text = '追加...'
$btnAdd.Location = New-Object Drawing.Point(640, 206); $btnAdd.Size = New-Object Drawing.Size(90, 27)
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
$btnDel.Location = New-Object Drawing.Point(640, 239); $btnDel.Size = New-Object Drawing.Size(90, 27)
$btnDel.Anchor = 'Top,Right'
$btnDel.Add_Click({
    foreach ($item in @($lst.SelectedItems)) { $lst.Items.Remove($item) }
    Save-Settings $txtZ.Text $txtE.Text.Trim() @($lst.Items)
})

$chk = New-Object Windows.Forms.CheckBox
$chk.Text = 'すでにFD製番・部品納期が入っている行も上書きする（納期回答は毎回最新に入れ直します）'
$chk.Checked = $Overwrite.IsPresent
$chk.Location = New-Object Drawing.Point(12, 316); $chk.AutoSize = $true

$btnRun = New-Object Windows.Forms.Button
$btnRun.Text = '実行'
$btnRun.Location = New-Object Drawing.Point(640, 312); $btnRun.Size = New-Object Drawing.Size(90, 32)
$btnRun.Anchor = 'Top,Right'

$txtLog = New-Object Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Location = New-Object Drawing.Point(12, 354); $txtLog.Size = New-Object Drawing.Size(718, 305)
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
        Save-Settings $txtZ.Text $txtE.Text.Trim() @($lst.Items)
        Invoke-Fill $txtZ.Text $txtO.Text.Trim() $txtE.Text.Trim() @($lst.Items) $chk.Checked $log
        [Windows.Forms.MessageBox]::Show($form, '完了しました。', '製番転記') | Out-Null
    } catch {
        & $log "エラー: $($_.Exception.Message)"
        [Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'エラー', 'OK', 'Error') | Out-Null
    } finally {
        $form.Cursor = 'Default'
        $btnRun.Enabled = $true
    }
})

$form.Controls.AddRange(@($lbl1, $txtZ, $btnZ, $lblO, $txtO, $btnO, $lblE, $txtE, $btnE, $lbl2, $lst, $btnAdd, $btnDel, $chk, $btnRun, $txtLog))
$form.Add_FormClosing({ try { Save-Settings $txtZ.Text $txtE.Text.Trim() @($lst.Items) } catch { } })
[void]$form.ShowDialog()

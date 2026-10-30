# 手配集計ツール
# 新規受注CSV（Shift-JIS）から、条件で絞り込んだ行を品目番号ごと・納期の近いものでまとめ、Excelに出力する。
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------- 集計処理 ----------------
function Read-OrderCsv([string]$path) {
  # Excelで開いたままでも読めるよう共有モードで開く
  $fs = [System.IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
  try {
    $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::GetEncoding(932))
    $text = $sr.ReadToEnd()
  } finally { $fs.Close() }
  $lines = @($text -split "`r?`n" | Where-Object { $_ -ne '' })
  if ($lines.Count -lt 2) { throw 'CSVにデータ行がありません。' }
  $header = $lines[0].Split(',')
  $rows = New-Object System.Collections.ArrayList
  for ($i = 1; $i -lt $lines.Count; $i++) { [void]$rows.Add($lines[$i].Split(',')) }
  return @{ Header = $header; Rows = $rows }
}

function Get-Col($header, [string]$name) {
  $i = [array]::IndexOf($header, $name)
  if ($i -lt 0) { throw "CSVに「$name」の列が見つかりません。" }
  return $i
}

function Test-FileLocked([string]$path) {
  try { $f = [System.IO.File]::Open($path, 'Open', 'ReadWrite', 'None'); $f.Close(); return $false } catch { return $true }
}

function D([string]$s) { [datetime]::ParseExact($s.Trim(), 'yyyyMMdd', $null) }

function Invoke-Shukei([string]$src, [string]$out, [string]$userKey, [bool]$excludeJit, [int]$win, [scriptblock]$log) {
  & $log "CSVを読み込み中: $src"
  $csv = Read-OrderCsv $src
  $h = $csv.Header
  $IX = @{}
  foreach ($n in 'オーダ番号','オーダ番号チェックデジット','納期','発注数','引当月（生産年月）','品目番号','図番','リビジョン','品名','規格･寸法',
                 '材料分類名','代表機種名','納入指示名','受入場所名','納入場所','単位名','改訂番号','ユーザー欄１','ユーザー欄２','EDI発注管理番号') {
    $IX[$n] = Get-Col $h $n
  }
  & $log "全 $($csv.Rows.Count) 行"

  $target = New-Object System.Collections.ArrayList
  foreach ($r in $csv.Rows) {
    if ($r[$IX['ユーザー欄１']].Trim() -ne $userKey) { continue }
    if ($excludeJit -and $r[$IX['納入指示名']].Trim() -eq 'ＪＩＴ') { continue }
    [void]$target.Add($r)
  }
  & $log "絞り込み後: $($target.Count) 行（ユーザー欄１ = 「$userKey」$(if($excludeJit){'、JIT除外'})）"
  if ($target.Count -eq 0) { throw '条件に合う行がありません。' }

  # 品目ごとに納期順に並べ、先頭の納期から $win 日以内を1グループにまとめる
  $groups = New-Object System.Collections.ArrayList
  foreach ($g in ($target | Group-Object { $_[$IX['品目番号']] })) {
    $sorted = @($g.Group | Sort-Object { $_[$IX['納期']] }, { $_[$IX['オーダ番号']] })
    $cur = $null; $start = $null
    foreach ($r in $sorted) {
      $d = D $r[$IX['納期']]
      if ($null -eq $start -or ($d - $start).Days -gt $win) {
        if ($cur) { [void]$groups.Add($cur) }
        $cur = New-Object System.Collections.ArrayList; $start = $d
      }
      [void]$cur.Add($r)
    }
    if ($cur) { [void]$groups.Add($cur) }
  }
  $groups = @($groups | Sort-Object { $_[0][$IX['納期']] }, { $_[0][$IX['品目番号']] })
  & $log "まとめた後: $($groups.Count) 行（$win 日以内をまとめる）"

  function U($g, $col) { (@($g | ForEach-Object { $_[$col].Trim() } | Where-Object { $_ -ne '' } | Sort-Object -Unique)) -join ' / ' }

  $hdr = @('No','品目番号','図番','リビジョン','改訂番号','品名','規格・寸法','材料分類名','手配納期(最早)','納期(最遅)','発注数合計','単位','元行数',
           '受入場所名','納入場所','ユーザー欄２','代表機種名','オーダ番号','要確認')
  $sum = New-Object 'object[,]' ($groups.Count + 1), $hdr.Count
  for ($c = 0; $c -lt $hdr.Count; $c++) { $sum[0, $c] = $hdr[$c] }
  $dh = @('集計No','品目番号','オーダ番号','CD','図番','品名','納期','発注数','受入場所名','納入場所','ユーザー欄２','代表機種名','引当月','EDI発注管理番号')
  $det = New-Object 'object[,]' ($target.Count + 1), $dh.Count
  for ($c = 0; $c -lt $dh.Count; $c++) { $det[0, $c] = $dh[$c] }

  $n = 0; $dr = 0; $warnRows = @(); $total = 0
  foreach ($g in $groups) {
    $n++
    $f = $g[0]
    $dates = @($g | ForEach-Object { $_[$IX['納期']] } | Sort-Object)
    $qty = 0; foreach ($r in $g) { $qty += [int]$r[$IX['発注数']] }
    $total += $qty
    $warn = @()
    foreach ($k in '図番','リビジョン','改訂番号','受入場所名','納入場所') {
      if (@($g | ForEach-Object { $_[$IX[$k]].Trim() } | Sort-Object -Unique).Count -gt 1) { $warn += "${k}混在" }
    }
    if (@($g | Where-Object { $_[$IX['引当月（生産年月）']].Trim() -eq '' }).Count -gt 0) { $warn += '引当月空欄あり' }
    if ($warn.Count) { $warnRows += $n + 1 }
    $vals = @($n, $f[$IX['品目番号']], (U $g $IX['図番']), (U $g $IX['リビジョン']), (U $g $IX['改訂番号']), $f[$IX['品名']].Trim(),
      (U $g $IX['規格･寸法']), (U $g $IX['材料分類名']), (D $dates[0]).ToOADate(), (D $dates[-1]).ToOADate(), $qty,
      $f[$IX['単位名']].Trim(), $g.Count, (U $g $IX['受入場所名']), (U $g $IX['納入場所']), (U $g $IX['ユーザー欄２']), (U $g $IX['代表機種名']),
      ((@($g | ForEach-Object { $_[$IX['オーダ番号']] })) -join ', '), ($warn -join ' / '))
    for ($c = 0; $c -lt $vals.Count; $c++) { $sum[$n, $c] = $vals[$c] }
    foreach ($r in $g) {
      $dr++
      $dv = @($n, $r[$IX['品目番号']], $r[$IX['オーダ番号']], $r[$IX['オーダ番号チェックデジット']], $r[$IX['図番']], $r[$IX['品名']].Trim(),
        (D $r[$IX['納期']]).ToOADate(), [int]$r[$IX['発注数']], $r[$IX['受入場所名']], $r[$IX['納入場所']].Trim(), $r[$IX['ユーザー欄２']].Trim(),
        $r[$IX['代表機種名']].Trim(), $r[$IX['引当月（生産年月）']].Trim(), $r[$IX['EDI発注管理番号']])
      for ($c = 0; $c -lt $dv.Count; $c++) { $det[$dr, $c] = $dv[$c] }
    }
  }

  & $log 'Excelに書き出し中...'
  $xl = $null; $wb = $null
  try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    while ($wb.Worksheets.Count -gt 2) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $ws = $wb.Worksheets.Item(1); $ws.Name = '手配集計'
    $wd = $wb.Worksheets.Item(2); $wd.Name = '明細'
    $ws.Range($ws.Cells(1, 1), $ws.Cells($groups.Count + 1, $hdr.Count)).Value2 = $sum
    $wd.Range($wd.Cells(1, 1), $wd.Cells($target.Count + 1, $dh.Count)).Value2 = $det
    $ws.Range($ws.Cells(2, 9), $ws.Cells($groups.Count + 1, 10)).NumberFormat = 'yyyy/mm/dd'
    $wd.Range($wd.Cells(2, 7), $wd.Cells($target.Count + 1, 7)).NumberFormat = 'yyyy/mm/dd'
    foreach ($sh in @($ws, $wd)) {
      $sh.Rows(1).Font.Bold = $true
      $sh.Rows(1).Interior.Color = 0xEBDDD9
      [void]$sh.UsedRange.AutoFilter()
      [void]$sh.UsedRange.Columns.AutoFit()
      $sh.Activate(); $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.FreezePanes = $true
    }
    $ws.Columns(18).ColumnWidth = 40
    foreach ($row in $warnRows) { $ws.Range($ws.Cells($row, 1), $ws.Cells($row, $hdr.Count)).Interior.Color = 0x99FFFF }
    $ws.Activate(); [void]$ws.Cells(1, 1).Select()
    if (Test-Path $out) { Remove-Item $out -Force -ErrorAction Stop }
    $wb.SaveAs($out, 51)
    $wb.Close($false); $wb = $null
  } finally {
    # 裏で起動したExcelが残らないよう確実に終了させる
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    $ws = $null; $wd = $null; $sh = $null; $wb = $null; $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
  }
  & $log "完了: 集計 $($groups.Count) 行 / 明細 $($target.Count) 行 / 発注数合計 $total / 要確認 $($warnRows.Count) 行"
  & $log "保存先: $out"
}

# ---------------- 画面 ----------------
$font = New-Object System.Drawing.Font('Meiryo UI', 10)
$form = New-Object System.Windows.Forms.Form
$form.Text = '手配集計ツール'
$form.Size = New-Object System.Drawing.Size(640, 520)
$form.StartPosition = 'CenterScreen'
$form.Font = $font
$form.AllowDrop = $true

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = '① 新規受注CSVを選ぶ（この画面にドラッグ＆ドロップでもOK）'
$lbl.Location = '16,16'; $lbl.Size = '600,22'
$form.Controls.Add($lbl)

$txt = New-Object System.Windows.Forms.TextBox
$txt.Location = '16,42'; $txt.Size = '480,26'
$form.Controls.Add($txt)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '参照...'
$btnBrowse.Location = '504,40'; $btnBrowse.Size = '100,30'
$form.Controls.Add($btnBrowse)

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = '② 条件'
$grp.Location = '16,80'; $grp.Size = '588,100'
$form.Controls.Add($grp)

$l1 = New-Object System.Windows.Forms.Label
$l1.Text = 'ユーザー欄１ が'; $l1.Location = '14,30'; $l1.Size = '110,22'
$grp.Controls.Add($l1)
$txtKey = New-Object System.Windows.Forms.TextBox
$txtKey.Text = 'J'; $txtKey.Location = '126,27'; $txtKey.Size = '60,26'
$grp.Controls.Add($txtKey)
$l1b = New-Object System.Windows.Forms.Label
$l1b.Text = 'と完全一致する行'; $l1b.Location = '192,30'; $l1b.Size = '140,22'
$grp.Controls.Add($l1b)

$chkJit = New-Object System.Windows.Forms.CheckBox
$chkJit.Text = '納入指示が JIT の行を除く'; $chkJit.Checked = $true
$chkJit.Location = '340,28'; $chkJit.Size = '240,24'
$grp.Controls.Add($chkJit)

$l2 = New-Object System.Windows.Forms.Label
$l2.Text = '品目番号ごとに、最初の納期から'; $l2.Location = '14,64'; $l2.Size = '220,22'
$grp.Controls.Add($l2)
$numWin = New-Object System.Windows.Forms.NumericUpDown
$numWin.Minimum = 0; $numWin.Maximum = 365; $numWin.Value = 7
$numWin.Location = '236,61'; $numWin.Size = '60,26'
$grp.Controls.Add($numWin)
$l2b = New-Object System.Windows.Forms.Label
$l2b.Text = '日以内の行をまとめる'; $l2b.Location = '302,64'; $l2b.Size = '200,22'
$grp.Controls.Add($l2b)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = '③ 集計してExcelを作る'
$btnRun.Location = '16,192'; $btnRun.Size = '588,44'
$btnRun.Font = New-Object System.Drawing.Font('Meiryo UI', 12, [System.Drawing.FontStyle]::Bold)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(217, 225, 242)
$form.Controls.Add($btnRun)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true; $logBox.ReadOnly = $true; $logBox.ScrollBars = 'Vertical'
$logBox.Location = '16,246'; $logBox.Size = '588,170'
$form.Controls.Add($logBox)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = 'できたExcelを開く'; $btnOpen.Enabled = $false
$btnOpen.Location = '16,426'; $btnOpen.Size = '290,36'
$form.Controls.Add($btnOpen)

$btnFolder = New-Object System.Windows.Forms.Button
$btnFolder.Text = '保存フォルダを開く'; $btnFolder.Enabled = $false
$btnFolder.Location = '314,426'; $btnFolder.Size = '290,36'
$form.Controls.Add($btnFolder)

$script:outPath = $null
$writeLog = {
  param($m)
  $logBox.AppendText("$m`r`n")
  [System.Windows.Forms.Application]::DoEvents()
}

$btnBrowse.Add_Click({
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = 'CSVファイル (*.csv)|*.csv|すべてのファイル (*.*)|*.*'
  $dlg.Title = '新規受注CSVを選択'
  $def = Join-Path $PSScriptRoot 'TEST'
  if (Test-Path $def) { $dlg.InitialDirectory = $def }
  if ($dlg.ShowDialog() -eq 'OK') { $txt.Text = $dlg.FileName }
})

$form.Add_DragEnter({ if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = 'Copy' } })
$form.Add_DragDrop({ $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop); if ($files) { $txt.Text = $files[0] } })

$btnRun.Add_Click({
  $src = $txt.Text.Trim('"', ' ')
  if (-not $src -or -not (Test-Path $src)) {
    [void][System.Windows.Forms.MessageBox]::Show('CSVファイルを選んでください。', '手配集計ツール'); return
  }
  $key = $txtKey.Text.Trim()
  if (-not $key) { [void][System.Windows.Forms.MessageBox]::Show('ユーザー欄１の値を入れてください。', '手配集計ツール'); return }
  # 出力名: CSV名の日付(8桁)を使う。なければ今日の日付
  $name = [System.IO.Path]::GetFileName($src)
  $m = [regex]::Match($name, '\d{8}')
  $stamp = if ($m.Success) { $m.Value } else { (Get-Date).ToString('yyyyMMdd') }
  $out = Join-Path (Split-Path $src -Parent) "手配集計($stamp).xlsx"
  if (Test-Path $out) {
    if (Test-FileLocked $out) {
      # Excelで開いたままだと上書きできないので、別名保存か中止を選んでもらう
      $alt = Join-Path (Split-Path $src -Parent) "手配集計($stamp)_$((Get-Date).ToString('HHmmss')).xlsx"
      $ans = [System.Windows.Forms.MessageBox]::Show(
        "「手配集計($stamp).xlsx」がExcelで開かれているため上書きできません。`r`n`r`n" +
        "はい　→ 別の名前（$([System.IO.Path]::GetFileName($alt))）で保存する`r`n" +
        "いいえ → 中止する（Excelでファイルを閉じてから、もう一度押してください）",
        '手配集計ツール', 'YesNo', 'Warning')
      if ($ans -ne 'Yes') { return }
      $out = $alt
    } else {
      $ans = [System.Windows.Forms.MessageBox]::Show("「手配集計($stamp).xlsx」はすでにあります。上書きしますか？", '手配集計ツール', 'YesNo', 'Question')
      if ($ans -ne 'Yes') { return }
    }
  }
  $logBox.Clear(); $btnRun.Enabled = $false; $btnOpen.Enabled = $false; $btnFolder.Enabled = $false
  $form.Cursor = 'WaitCursor'
  try {
    Invoke-Shukei $src $out $key $chkJit.Checked ([int]$numWin.Value) $writeLog
    $script:outPath = $out
    $btnOpen.Enabled = $true; $btnFolder.Enabled = $true
  } catch {
    & $writeLog "エラー: $($_.Exception.Message)"
    [void][System.Windows.Forms.MessageBox]::Show("エラーが発生しました。`r`n$($_.Exception.Message)", '手配集計ツール', 'OK', 'Error')
  } finally {
    $form.Cursor = 'Default'; $btnRun.Enabled = $true
  }
})

$btnOpen.Add_Click({ if ($script:outPath -and (Test-Path $script:outPath)) { Start-Process $script:outPath } })
$btnFolder.Add_Click({ if ($script:outPath) { Start-Process explorer.exe "/select,`"$script:outPath`"" } })

# テスト用: -Src/-Out を渡すと画面なしで実行
if ($args.Count -ge 2) {
  Invoke-Shukei $args[0] $args[1] 'J' $true 7 $writeLog
  Write-Output $logBox.Text
  return
}
[void]$form.ShowDialog()

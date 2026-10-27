param(
    [Parameter(Mandatory=$true)][string]$InputDocx,
    [Parameter(Mandatory=$true)][string]$Manifest,
    [Parameter(Mandatory=$true)][string]$OutputDocx,
    [string]$MathTypeTemplate = 'C:\Program Files (x86)\MathType\Office Support\64\MathType Commands 2016.dotm'
)
$ErrorActionPreference = 'Stop'

# 只处理显式指定的占位符；不自动转换现有公式。
$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
$manifestPath = (Resolve-Path -LiteralPath $Manifest).Path
$outputPath = [IO.Path]::GetFullPath($OutputDocx)
if ([IO.Path]::GetExtension($inputPath) -ine '.docx' -or [IO.Path]::GetExtension($outputPath) -ine '.docx') { throw '输入和输出必须为 docx。' }
if ($inputPath -ieq $outputPath) { throw '输出不能与输入相同。' }
if (Test-Path -LiteralPath $outputPath) { throw '输出已存在；请指定新的输出文件名。' }
if (Test-Path -LiteralPath ($outputPath + '.mathtype.json')) { throw '输出报告已存在；请指定新的输出文件名。' }
if (-not (Test-Path -LiteralPath (Split-Path -Parent $outputPath) -PathType Container)) { throw '输出目录不存在。' }
if (-not (Test-Path -LiteralPath $MathTypeTemplate -PathType Leaf)) { throw 'MathType Word 模板不存在，请检查路径和 Word 位数。' }
if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) { throw '检测到 Word 正在运行；请保存并关闭已有会话后重试。' }
$items = @(Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json)
if ($items.Count -lt 1) { throw '公式清单为空。' }
$seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($item in $items) {
    if ($item.marker -cnotmatch '^\[\[MT:[A-Za-z0-9_-]{1,60}\]\]$') { throw "占位符格式无效: $($item.marker)" }
    if (-not $seen.Add([string]$item.marker)) { throw "清单含重复占位符: $($item.marker)" }
    if ([string]::IsNullOrWhiteSpace([string]$item.tex) -or $item.tex -match '[\r\n$]') { throw "TeX 为空、含换行或包含美元定界符: $($item.marker)" }
    if ($item.layout -notin @('inline','display')) { throw "layout 只能是 inline 或 display: $($item.marker)" }
}

$sourceHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash
$stagePath = Join-Path (Split-Path -Parent $outputPath) ('.mathtype-' + [guid]::NewGuid().ToString('N') + '.docx')
Copy-Item -LiteralPath $inputPath -Destination $stagePath
$word = $null
$doc = $null
$results = @()
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($stagePath, $false, $false)
    $null = $word.AddIns.Add($MathTypeTemplate, $true)
    $beforeFieldCodes = @($doc.Fields | ForEach-Object { $_.Code.Text } | Where-Object { $_ -notmatch '^\s*EMBED\b' })
    $beforeBookmarks = @($doc.Bookmarks | ForEach-Object { $_.Name } | Sort-Object)
    foreach ($item in $items) {
        $range = $doc.Content
        $find = $range.Find
        $find.ClearFormatting()
        $find.Text = [string]$item.marker
        $find.Forward = $true
        $find.Wrap = 0
        $find.MatchWildcards = $false
        $find.MatchCase = $true
        $find.MatchWholeWord = $false
        if (-not $find.Execute()) { throw "未找到占位符: $($item.marker)" }
        $start = $range.Start
        $duplicate = $doc.Range($range.End, $doc.Content.End - 1)
        $next = $duplicate.Find
        $next.ClearFormatting(); $next.Text = [string]$item.marker; $next.Forward = $true; $next.Wrap = 0; $next.MatchWildcards = $false; $next.MatchCase = $true; $next.MatchWholeWord = $false
        if ($next.Execute()) { throw "文档含重复占位符: $($item.marker)" }
        if ($range.Fields.Count -gt 0 -or $range.Hyperlinks.Count -gt 0 -or $range.Bookmarks.Count -gt 0) { throw "占位符与域、链接或书签重叠: $($item.marker)" }
        # 两种布局都仅转换占位符；独立公式的段落、制表位和编号由文档模板控制。
        $delimited = '$' + $item.tex + '$'
        $range.Text = $delimited
        $doc.Range($start, $start + $delimited.Length).Select()
        $countBefore = $doc.InlineShapes.Count
        $word.Run('MTCommand_TeXToggle')
        if ($doc.InlineShapes.Count -ne $countBefore + 1) { throw "公式对象数异常: $($item.marker)" }
        # OLE 对象范围可包含 EMBED 域；独立公式的选择区可能额外包含换行。
        $selectionStart = $word.Selection.Start
        $selectionEnd = $word.Selection.End
        if ($selectionStart -ne $start -or $selectionEnd -le $start) { throw "转换后的选择区异常: $($item.marker)" }
        $candidates = @($doc.InlineShapes | Where-Object { $_.Range.Start -ge $selectionStart -and $_.Range.End -le $selectionEnd })
        if ($candidates.Count -ne 1) { throw "生成对象未处于预期选择区: $($item.marker)" }
        $object = $candidates[0]
        if ($object.Type -ne 1 -or $object.OLEFormat.ProgID -ne 'Equation.DSMT4' -or $object.Width -lt 1 -or $object.Height -lt 1) { throw "生成对象无效: $($item.marker)" }
        $results += [ordered]@{ marker=$item.marker; layout=$item.layout; tex=$item.tex; progid=$object.OLEFormat.ProgID; width_pt=$object.Width; height_pt=$object.Height }
        Write-Output "Converted $($item.marker)"
    }
    $afterFieldCodes = @($doc.Fields | ForEach-Object { $_.Code.Text } | Where-Object { $_ -notmatch '^\s*EMBED\b' })
    $afterBookmarks = @($doc.Bookmarks | ForEach-Object { $_.Name } | Sort-Object)
    if (($beforeFieldCodes | ConvertTo-Json -Compress) -cne ($afterFieldCodes | ConvertTo-Json -Compress)) { throw '域代码发生变化，停止发布。' }
    if (($beforeBookmarks | ConvertTo-Json -Compress) -cne ($afterBookmarks | ConvertTo-Json -Compress)) { throw '书签列表发生变化，停止发布。' }
    $doc.Save()
    $doc.Close(0); $doc = $null
    $word.Quit(); $word = $null
    if ((Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash -ne $sourceHash) { throw '原文件哈希发生变化，停止发布。' }
    [IO.File]::Move($stagePath, $outputPath)
    $report = [ordered]@{ source=$inputPath; source_sha256=$sourceHash; manifest=$manifestPath; manifest_sha256=(Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash; output=$outputPath; output_sha256=(Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash; formulas=$results; note='对象检查通过；公式数学含义、排版、编号和引用必须人工复核。' }
    $reportPath = $outputPath + '.mathtype.json'
    $reportText = $report | ConvertTo-Json -Depth 8
    # 使用 CreateNew 防止意外覆盖同名报告；若报告写入失败，保留已发布的输出并报错。
    $reportStream = [IO.File]::Open($reportPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $writer = [IO.StreamWriter]::new($reportStream, [Text.UTF8Encoding]::new($false))
        try { $writer.Write($reportText) } finally { $writer.Dispose() }
    } finally { $reportStream.Dispose() }
    Write-Output "Created $outputPath"
} finally {
    if ($null -ne $doc) { try { $doc.Close(0) } catch {} }
    if ($null -ne $word) { try { $word.Quit() } catch {} }
    if (Test-Path -LiteralPath $stagePath -PathType Leaf) { Remove-Item -LiteralPath $stagePath -Force }
}

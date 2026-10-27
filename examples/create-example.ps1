param([Parameter(Mandatory=$true)][string]$OutputDocx)
$ErrorActionPreference = 'Stop'
$output = [IO.Path]::GetFullPath($OutputDocx)
if ([IO.Path]::GetExtension($output) -ine '.docx') { throw '输出必须是 docx。' }
if (Test-Path -LiteralPath $output) { throw '输出已存在，请指定新文件名。' }
if (-not (Test-Path -LiteralPath (Split-Path -Parent $output) -PathType Container)) { throw '输出目录不存在。' }
if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) { throw '请保存并关闭已打开的 Word。' }
$word=$null; $doc=$null; $createdWordId=$null
try {
    $word=New-Object -ComObject Word.Application
    $created = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
    if ($created.Count -eq 1) { $createdWordId = $created[0].Id }
    $word.Visible=$false; $word.DisplayAlerts=0
    $doc=$word.Documents.Add()
    $range=$doc.Range(0,0)
    $range.InsertAfter("MathType 示例（虚构数据）`r行内：[[MT:inline]]`r分式与求和：`r`t[[MT:sum]]`t（[[SEQFIELD]]）`r矩阵：`r`t[[MT:matrix]]`r向量：[[MT:vector]]`r根式：`r`t[[MT:root]]`r交叉引用：式（[[REFFIELD]]）`r")
    $doc.Content.ParagraphFormat.LineSpacingRule=3
    $doc.Content.ParagraphFormat.LineSpacing=42
    foreach ($marker in @('[[MT:sum]]','[[MT:matrix]]','[[MT:root]]')) {
        $range=$doc.Content; $find=$range.Find
        $find.Text=$marker; $find.Forward=$true; $find.Wrap=0; $find.MatchWildcards=$false
        if (-not $find.Execute()) { throw "找不到 $marker" }
        $range.ParagraphFormat.TabStops.ClearAll()
        $null=$range.ParagraphFormat.TabStops.Add(212.6,1)
        $null=$range.ParagraphFormat.TabStops.Add(425.2,2)
    }
    $range=$doc.Content; $find=$range.Find; $find.Text='[[SEQFIELD]]'; $find.Forward=$true; $find.Wrap=0; $find.MatchWildcards=$false
    if (-not $find.Execute()) { throw '找不到自动编号标记。' }
    $range.Text=''; $field=$doc.Fields.Add($range,-1,'SEQ sampleeq \* ARABIC',$true); $null=$field.Update()
    $null=$doc.Bookmarks.Add('sample_eq1',$field.Result)
    $range=$doc.Content; $find=$range.Find; $find.Text='[[REFFIELD]]'; $find.Forward=$true; $find.Wrap=0; $find.MatchWildcards=$false
    if (-not $find.Execute()) { throw '找不到交叉引用标记。' }
    $range.Text=''; $field=$doc.Fields.Add($range,-1,'REF sample_eq1 \h',$true); $null=$field.Update()
    $doc.SaveAs2($output,16)
    Write-Output "Created example: $output"
} finally {
    if ($null -ne $doc) { $doc.Close(0) }
    if ($null -ne $word) { $word.Quit() }
    $field=$null; $find=$null; $range=$null; $doc=$null; $word=$null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # 等待退出，不终止进程；若另有 Word 打开，转换脚本仍会拒绝运行。
    if ($null -ne $createdWordId) { try { Wait-Process -Id $createdWordId -Timeout 15 -ErrorAction Stop } catch {} }
}

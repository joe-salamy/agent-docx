$ErrorActionPreference = 'Stop'
$inputJson = [Console]::In.ReadToEnd() | ConvertFrom-Json
$mutex = New-Object Threading.Mutex($false, 'Local\agent-docx-word')
$acquired = $false
$word = $null
$doc = $null
$temp = $null
try {
  $acquired = $mutex.WaitOne([TimeSpan]::FromSeconds(120))
  if (-not $acquired) { throw 'Word renderer mutex timeout' }
  $temp = Join-Path ([IO.Path]::GetTempPath()) ('agent-docx-' + [Guid]::NewGuid().ToString('N'))
  [IO.Directory]::CreateDirectory($temp) | Out-Null
  $path = Join-Path $temp 'render.docx'
  [IO.File]::WriteAllBytes($path, [Convert]::FromBase64String([string]$inputJson.docxBase64))
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.ScreenUpdating = $false
  $oldSecurity = $word.AutomationSecurity
  $word.AutomationSecurity = 3
  $availableFonts = @($word.FontNames)
  foreach ($family in @($inputJson.requestedFontFamilies)) {
    if ($availableFonts -notcontains [string]$family) { throw ('Requested font is not installed in Word: ' + [string]$family) }
  }
  [Console]::Out.WriteLine((@{kind='started'; hwnd=$word.Hwnd} | ConvertTo-Json -Compress))
  $doc = $word.Documents.Open($path, $false, $true, $false)
  $doc.Repaginate()
  $pages = $doc.ComputeStatistics(2)
  $totalLines = $doc.ComputeStatistics(1)
  $pageLines = @()
  for ($page = 1; $page -le $pages; $page++) {
    $pageStart = $doc.GoTo(1, 1, $page)
    if ($page -lt $pages) {
      $nextPage = $doc.GoTo(1, 1, $page + 1)
      $pageRange = $doc.Range($pageStart.Start, $nextPage.Start)
      [Runtime.InteropServices.Marshal]::FinalReleaseComObject($nextPage) | Out-Null
    } else {
      $pageRange = $doc.Range($pageStart.Start, $doc.Content.End)
    }
    $pageLines += $pageRange.ComputeStatistics(1)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($pageRange) | Out-Null
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($pageStart) | Out-Null
  }
  $docEnd = $doc.Content.End
  $last = $doc.GoTo(1, 1, $pages)
  $range = $doc.Range($last.Start, $docEnd)
  $lines = $range.ComputeStatistics(1)
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) | Out-Null
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($last) | Out-Null
  [Console]::Out.WriteLine((@{kind='result'; pageCount=$pages; totalBodyLines=$totalLines; bodyLinesByPage=$pageLines; bodyLinesOnLastPage=$lines; version=$word.Version; build=$word.Build; activePrinter=$word.ActivePrinter} | ConvertTo-Json -Compress))
} catch {
  [Console]::Out.WriteLine((@{kind='error'; message=$_.Exception.Message} | ConvertTo-Json -Compress))
  exit 1
} finally {
  if ($doc -ne $null) { $doc.Close(0); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null }
  if ($word -ne $null) { if ($oldSecurity -ne $null) { $word.AutomationSecurity = $oldSecurity }; $word.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null }
  if ($temp -ne $null -and [IO.Directory]::Exists($temp)) { [IO.Directory]::Delete($temp, $true) }
  if ($acquired) { $mutex.ReleaseMutex() }
  $mutex.Dispose()
}

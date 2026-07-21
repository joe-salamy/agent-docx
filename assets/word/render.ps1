$ErrorActionPreference = 'Stop'
$inputJson = [Console]::In.ReadToEnd() | ConvertFrom-Json
$mutex = New-Object Threading.Mutex($false, 'Local\md-page-count-word')
$acquired = $false
$word = $null
$doc = $null
$temp = $null
try {
  $acquired = $mutex.WaitOne([TimeSpan]::FromSeconds(120))
  if (-not $acquired) { throw 'Word renderer mutex timeout' }
  $temp = Join-Path ([IO.Path]::GetTempPath()) ('md-page-count-' + [Guid]::NewGuid().ToString('N'))
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
  $docEnd = $doc.Content.End
  $last = $doc.GoTo(1, 1, $pages)
  $range = $doc.Range($last.Start, $docEnd)
  $lines = $range.ComputeStatistics(1)
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) | Out-Null
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($last) | Out-Null
  [Console]::Out.WriteLine((@{kind='result'; pageCount=$pages; bodyLinesOnLastPage=$lines; version=$word.Version; build=$word.Build; activePrinter=$word.ActivePrinter} | ConvertTo-Json -Compress))
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

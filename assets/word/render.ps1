$ErrorActionPreference = 'Stop'
$utf8 = New-Object Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
$inputJson = [Console]::In.ReadToEnd() | ConvertFrom-Json
function Write-Frame($value) {
  [Console]::Out.WriteLine(($value | ConvertTo-Json -Compress -Depth 5))
}
if ($inputJson.protocolVersion -ne 2) {
  Write-Frame @{kind='error'; protocolVersion=2; message='Unsupported protocol version'}
  exit 1
}
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
  Write-Frame @{kind='started'; protocolVersion=2; hwnd=$word.Hwnd}
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
  Write-Frame @{
    kind='summary'
    protocolVersion=2
    pageCount=$pages
    totalBodyLines=$totalLines
    bodyLinesByPage=@($pageLines)
    bodyLinesOnLastPage=$lines
    version=[string]$word.Version
    build=[string]$word.Build
    activePrinter=[string]$word.ActivePrinter
  }
  try {
    $paragraphValues = @()
    foreach ($paragraphIdValue in @($inputJson.paragraphIds)) {
      $paragraphId = [string]$paragraphIdValue
      $bookmark = $null
      $paragraphRange = $null
      $startProbe = $null
      $endProbe = $null
      $finalRange = $null
      try {
        if (-not $doc.Bookmarks.Exists($paragraphId)) { throw ('Bookmark not found: ' + $paragraphId) }
        $bookmark = $doc.Bookmarks.Item($paragraphId)
        $paragraphRange = $bookmark.Range.Duplicate
        $lineCount = [int]$paragraphRange.ComputeStatistics(1)
        if ($lineCount -lt 1) { throw ('Invalid line count for bookmark: ' + $paragraphId) }
        $startProbe = $paragraphRange.Duplicate
        $startProbe.Collapse(1)
        $startPage = [int]$startProbe.Information(3)
        $endProbe = $paragraphRange.Duplicate
        $endProbe.Collapse(0)
        $endPage = [int]$endProbe.Information(3)
        $low = [int]$paragraphRange.Start
        $high = [int]$paragraphRange.End
        while ($low -lt $high) {
          $middle = [int][Math]::Floor(($low + $high) / 2)
          $suffix = $doc.Range($middle, $paragraphRange.End)
          try {
            $suffixLines = [int]$suffix.ComputeStatistics(1)
          } finally {
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($suffix) | Out-Null
          }
          if ($suffixLines -le 1) { $high = $middle } else { $low = $middle + 1 }
        }
        $finalRange = $doc.Range($low, $paragraphRange.End)
        $finalText = [string]$finalRange.Text
        if ($finalText.EndsWith([string][char]13)) {
          $finalText = $finalText.Substring(0, $finalText.Length - 1)
        }
        $footnoteRecords = @()
        for ($footnoteIndex = 1; $footnoteIndex -le $doc.Footnotes.Count; $footnoteIndex++) {
          $footnote = $null
          $reference = $null
          try {
            $footnote = $doc.Footnotes.Item($footnoteIndex)
            $reference = $footnote.Reference
            if ($reference.Start -ge $finalRange.Start -and $reference.Start -lt $finalRange.End) {
              $footnoteRecords += [PSCustomObject]@{Position=[int]$reference.Start; Index=[int]$footnote.Index}
            }
          } finally {
            if ($reference -ne $null) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference) | Out-Null }
            if ($footnote -ne $null) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($footnote) | Out-Null }
          }
        }
        $footnoteRecords = @($footnoteRecords | Sort-Object Position)
        $footnoteCursor = 0
        $builder = New-Object Text.StringBuilder
        foreach ($character in $finalText.ToCharArray()) {
          if ([int]$character -eq 2) {
            if ($footnoteCursor -ge $footnoteRecords.Count) { throw 'Unmatched footnote reference control character' }
            [void]$builder.Append([string]$footnoteRecords[$footnoteCursor].Index)
            $footnoteCursor++
          } else {
            [void]$builder.Append($character)
          }
        }
        if ($footnoteCursor -ne $footnoteRecords.Count) { throw 'Unmatched footnote reference range' }
        $paragraphValues += [PSCustomObject]@{
          id=$paragraphId
          lineCount=$lineCount
          startPage=$startPage
          endPage=$endPage
          finalLineText=$builder.ToString()
        }
      } finally {
        if ($finalRange -ne $null) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($finalRange) | Out-Null }
        if ($endProbe -ne $null) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($endProbe) | Out-Null }
        if ($startProbe -ne $null) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($startProbe) | Out-Null }
        if ($paragraphRange -ne $null) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($paragraphRange) | Out-Null }
        if ($bookmark -ne $null) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($bookmark) | Out-Null }
      }
    }
    Write-Frame @{kind='paragraphs'; protocolVersion=2; status='ok'; value=@($paragraphValues)}
  } catch {
    Write-Frame @{kind='paragraphs'; protocolVersion=2; status='error'; message=$_.Exception.Message}
  }
} catch {
  Write-Frame @{kind='error'; protocolVersion=2; message=$_.Exception.Message}
  exit 1
} finally {
  if ($doc -ne $null) { $doc.Close(0); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null }
  if ($word -ne $null) { if ($oldSecurity -ne $null) { $word.AutomationSecurity = $oldSecurity }; $word.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null }
  if ($temp -ne $null -and [IO.Directory]::Exists($temp)) { [IO.Directory]::Delete($temp, $true) }
  if ($acquired) { $mutex.ReleaseMutex() }
  $mutex.Dispose()
}

param(
  [string]$ExcelPath = "C:\Users\User\Downloads\物件總表 - 複製.xlsm",
  [string]$JsonPath = "C:\Users\User\Downloads\物件總表_2026-08-02_車位及土地欄位完整修正.json",
  [string]$OutputCsv = "C:\Users\User\Downloads\物件總表_五欄核對差異.csv"
)

$ErrorActionPreference = "Stop"
function Clean([object]$Value) {
  $text = ([string]$Value).Trim() -replace '\s+', '' -replace '％','%' -replace '／','/'
  $text = $text -replace '^(臨路|臨|面寬|面|深度|深)', '' -replace '(公尺|米)$',''
  if ($text -match '^[-—–－_無未填沒有]$') { return '' }
  return $text
}
function Same([object]$A, [object]$B, [bool]$ZeroAsBlank = $false) {
  $left = Clean $A; $right = Clean $B
  if ($ZeroAsBlank) { if ($left -match '^0(?:\.0+)?$') { $left = '' }; if ($right -match '^0(?:\.0+)?$') { $right = '' } }
  return $left -eq $right
}

$document = Get-Content -Raw -Encoding UTF8 -LiteralPath $JsonPath | ConvertFrom-Json
$records = @($document.records | Where-Object { -not $_.archived -and $_.status -eq '委託中' })
$byNo = @{}; foreach ($record in $records) { if ($record.propertyNo) { $byNo[[string]$record.propertyNo.Trim().ToUpper()] = $record } }
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$differences = New-Object System.Collections.Generic.List[object]
$matched = 0; $missing = 0
try {
  $book = $excel.Workbooks.Open($ExcelPath, 0, $true)
  $sheet = $book.Worksheets.Item('物件總表')
  for ($row = 2; $row -le $sheet.UsedRange.Rows.Count; $row++) {
    $propertyNo = ([string]$sheet.Cells.Item($row, 1).Text).Trim().ToUpper()
    if ($propertyNo -notmatch '^[A-Z]{2}\d+') { continue }
    $record = $byNo[$propertyNo]
    if (-not $record) { $missing++; continue }
    $matched++
    $road = [string]$sheet.Cells.Item($row, 5).Text
    $dimensions = [string]$sheet.Cells.Item($row, 6).Text
    $frontage = if ($dimensions -match '面(?:寬)?\s*([^\r\n]+)') { $Matches[1] } else { '' }
    $depth = if ($dimensions -match '深(?:度)?\s*([^\r\n]+)') { $Matches[1] } else { '' }
    $zoning = [string]$sheet.Cells.Item($row, 7).Text
    $coverageFar = [string]$sheet.Cells.Item($row, 8).Text
    $parts = @((Clean $coverageFar) -split '/')
    $coverage = if ($parts.Count -gt 0) { $parts[0] } else { '' }
    $far = if ($parts.Count -gt 1) { $parts[1] } else { '' }
    $isHouse = [string]$record.type -notmatch '土地'
    $checks = @(
      @{ Field='臨路'; Excel=$road; App=$record.road; Zero=$false },
      @{ Field='面寬'; Excel=$frontage; App=$record.frontage; Zero=$isHouse },
      @{ Field='深度'; Excel=$depth; App=$record.depth; Zero=$isHouse },
      @{ Field='使用分區'; Excel=$zoning; App=$record.zoning; Zero=$false },
      @{ Field='建蔽率'; Excel=$coverage; App=$record.coverage; Zero=$false },
      @{ Field='容積率'; Excel=$far; App=$record.far; Zero=$false }
    )
    foreach ($check in $checks) {
      if (-not (Same $check.Excel $check.App $check.Zero)) {
        $differences.Add([pscustomobject]@{ 物件編號=$propertyNo; 案名=$record.caseName; 欄位=$check.Field; Excel正確內容=(Clean $check.Excel); 單機目前內容=(Clean $check.App) })
      }
    }
  }
}
finally {
  if ($book) { $book.Close($false) }
  $excel.Quit()
}
$differences | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
[pscustomobject]@{ 委託中=$records.Count; Excel配對=$matched; Excel找不到單機=$missing; 差異欄位=$differences.Count; 報告=$OutputCsv }

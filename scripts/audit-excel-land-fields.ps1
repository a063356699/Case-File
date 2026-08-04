param(
  [string]$JsonPath = "C:\Users\User\Downloads\物件總表_2026-08-01 (最後修改11581 0943).json",
  [string]$WorkbookRoot = "C:\Users\User\Downloads\售屋資料表+不動產說明書\售屋資料表+不動產說明書",
  [string]$OutputJson = "C:\Users\User\Downloads\物件總表_2026-08-01_0943_土地欄位核對修正.json",
  [string]$AuditCsv = "C:\Users\User\Downloads\物件總表_2026-08-01_0943_土地欄位核對報告.csv"
)

$ErrorActionPreference = "Stop"

function Normalize-Text([string]$Text) {
  if (-not $Text) { return "" }
  return (($Text -replace '[\s　\-—_｜|\[\]【】()（）,，、.。之號市縣區]', '') -replace '臺', '台').ToLower()
}

function Clean-Measure([object]$Value, [string]$Prefix) {
  $text = [string]$Value
  $text = $text.Trim() -replace '^(約|土地|臨路|臨|面寬|面|深度|深)\s*', '' -replace '\s*公?米\s*$', ''
  if ($Prefix -eq '臨' -and $text -match '^無.*路$') { return "" }
  if ($text -match '^[-－—–_／/、.．無未填沒有0]*$') { return "" }
  return $text.Trim()
}

function Cell-Text($Sheet, [string[]]$Labels) {
  foreach ($label in $Labels) {
    $found = $Sheet.Rows.Item(1).Find($label)
    if ($null -ne $found) {
      try { return [string]$Sheet.Cells.Item(2, $found.Column).Text }
      finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($found) }
    }
  }
  return $null
}

function Header-Values($Sheet, [string]$Pattern) {
  $values = @()
  $columnCount = [Math]::Max(1, [int]$Sheet.UsedRange.Columns.Count)
  for ($column = 1; $column -le $columnCount; $column++) {
    $header = ([string]$Sheet.Cells.Item(1, $column).Text).Trim()
    if ($header -match $Pattern) {
      $value = ([string]$Sheet.Cells.Item(2, $column).Text).Trim()
      if ($value) { $values += $value }
    }
  }
  return @($values)
}

$document = Get-Content -Raw -Encoding UTF8 -LiteralPath $JsonPath | ConvertFrom-Json
$active = @($document.records | Where-Object { -not $_.archived -and $_.status -eq '委託中' })
$files = @(Get-ChildItem -LiteralPath $WorkbookRoot -Recurse -File | Where-Object {
  $_.Extension -match '^\.xls[mx]?$' -and $_.Name -notlike '~$*'
} | ForEach-Object {
  [pscustomobject]@{ File = $_; Name = Normalize-Text $_.BaseName }
})

function Find-Workbook($Record) {
  $address = Normalize-Text $Record.address
  $caseName = Normalize-Text $Record.caseName
  $candidates = @($files | Where-Object {
    ($address.Length -ge 5 -and $_.Name.Contains($address)) -or
    ($caseName.Length -ge 6 -and $_.Name.Contains($caseName))
  })
  if (-not $candidates.Count) {
    $manualAddress = switch ($Record.propertyNo) {
      'EG0495413' { '安平路3047樓12' }
      'EG0507776' { '新義南路20' }
      'EG0522916' { '中華路一段323巷6' }
      default { '' }
    }
    if ($manualAddress) { $candidates = @($files | Where-Object { $_.Name.Contains($manualAddress) }) }
  }
  if (-not $candidates.Count) { return $null }
  return $candidates | Sort-Object @{Expression={
    $path = $_.File.FullName
    $penalty = 0
    if ($path -match '複製|檔案與文件|南北-其他區') { $penalty += 10000 }
    if ($_.File.Extension -ne '.xlsm') { $penalty += 1000 }
    $penalty + $path.Length
  }} | Select-Object -First 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$audit = New-Object System.Collections.Generic.List[object]
$matched = 0
$unmatched = New-Object System.Collections.Generic.List[object]
$noResponse = New-Object System.Collections.Generic.List[object]

try {
  foreach ($record in $active) {
    $candidate = Find-Workbook $record
    if ($null -eq $candidate) {
      $unmatched.Add([pscustomobject]@{ 物件編號=$record.propertyNo; 案名=$record.caseName; 地址=$record.address })
      continue
    }
    $workbook = $null
    $sheet = $null
    try {
      $workbook = $excel.Workbooks.Open($candidate.File.FullName, 0, $true)
      try { $sheet = $workbook.Worksheets.Item('回應') } catch { $sheet = $null }
      $parkingOwnership = [string]$record.parkingOwnership
      $parkingMethod = [string]$record.parkingMethod
      $parkingNo = [string]$record.parkingNo
      if ($null -eq $sheet) {
        try {
          $sheet = $workbook.Worksheets.Item('總表')
          $road = Clean-Measure ([string]$sheet.Range('Q2').Text) '臨'
          $pair = [string]$sheet.Range('R2').Text
          $frontage = if ($pair -match '面(?:寬)?\s*([^\r\n]+)') { Clean-Measure $Matches[1] '面' } else { '' }
          $depth = if ($pair -match '深(?:度)?\s*([^\r\n]+)') { Clean-Measure $Matches[1] '深' } else { '' }
          $zoningRaw = [string]$sheet.Range('S2').Text
          $coverageRaw = [string]$sheet.Range('T2').Text
        }
        catch {
          $noResponse.Add([pscustomobject]@{ 物件編號=$record.propertyNo; 案名=$record.caseName; Excel=$candidate.File.FullName })
          continue
        }
      } else {
        # 「回應」工作表固定欄位：Q臨路、R面寬、S深度、T使用分區、U建蔽率/容積率。
        # 直接讀取固定儲存格，避免 Excel 欄位標題有空格或文字變更時抓錯欄位。
        $road = Clean-Measure ([string]$sheet.Range('Q2').Text) '臨'
        $frontage = Clean-Measure ([string]$sheet.Range('R2').Text) '面'
        $depth = Clean-Measure ([string]$sheet.Range('S2').Text) '深'
        $zoningRaw = [string]$sheet.Range('T2').Text
        $coverageRaw = [string]$sheet.Range('U2').Text
        $parkingOwnershipRaw = (Header-Values $sheet '^車位$' | Select-Object -First 1)
        if (-not $parkingOwnershipRaw) { $parkingOwnershipRaw = ([string]$sheet.Range('AL2').Text).Trim() }
        $parkingOwnershipCompact = $parkingOwnershipRaw -replace '\s+', '' -replace '／', '+' -replace '/', '+'
        $parkingOwnership = switch -Regex ($parkingOwnershipCompact) {
          '無車位.*無產權' { '無車位'; break }
          '有車位.*主附建物內含.*停自有地' { '停自有地'; break }
          '有車位.*公設內含.*固定車位' { '固定車位'; break }
          '有車位.*獨立產權.*固定車位' { '固定車位'; break }
          '有車位.*無產權.*車位另租' { '車位另租'; break }
          '有車位.*無產權.*抽籤決定' { '抽籤決定'; break }
          default { $parkingOwnershipRaw }
        }
        $parkingTypeCells = @(Header-Values $sheet '^車位型態')
        if (-not $parkingTypeCells.Count) { $parkingTypeCells = @('AM2','AN2','AO2','AP2','AQ2') | ForEach-Object { ([string]$sheet.Range($_).Text).Trim() } | Where-Object { $_ } }
        $parkingMethod = ($parkingTypeCells | Where-Object { $_ -in @('坡道/平面','坡道/機械','昇降/平面','昇降/機械','庭院','車庫','平移/機械') } | Select-Object -First 1)
        if (-not $parkingMethod) { $parkingMethod = $parkingTypeCells | Select-Object -First 1 }
        $parkingNo = (Header-Values $sheet '^車位編號' | Select-Object -First 1)
        if (-not $parkingNo) { $parkingNo = ([string]$sheet.Range('AR2').Text).Trim() }
      }
      $zoning = ([string]$zoningRaw).Trim()
      if ($zoning -match '^[-－—–_／/、.．無未填沒有0]*$') { $zoning = '' }
      $coverageText = ([string]$coverageRaw).Trim() -replace '％', '%' -replace '\s+', ''
      if ($coverageText -match '^[-－—–_／/、.．無未填沒有0]*$') { $coverageText = '' }
      $parts = @($coverageText -split '[/／]')
      $coverage = if ($parts.Count -gt 0) { ($parts[0] -replace '%','').Trim() } else { '' }
      $far = if ($parts.Count -gt 1) { ($parts[1] -replace '%','').Trim() } else { '' }

      foreach ($change in @(
        @{Key='road'; Label='臨路'; New=$road},
        @{Key='frontage'; Label='面寬'; New=$frontage},
        @{Key='depth'; Label='深度'; New=$depth},
        @{Key='zoning'; Label='使用分區'; New=$zoning},
        @{Key='coverage'; Label='建蔽率'; New=$coverage},
        @{Key='far'; Label='容積率'; New=$far},
        @{Key='parkingOwnership'; Label='車位產權'; New=$parkingOwnership},
        @{Key='parkingMethod'; Label='車位型態'; New=$parkingMethod},
        @{Key='parkingNo'; Label='車位編號'; New=$parkingNo}
      )) {
        $old = [string]$record.($change.Key)
        $new = [string]$change.New
        if ($old.Trim() -ne $new.Trim()) {
          $audit.Add([pscustomobject]@{
            物件編號=$record.propertyNo; 案名=$record.caseName; 欄位=$change.Label
            原內容=$old; Excel正確內容=$new; Excel檔案=$candidate.File.Name
          })
          $record.($change.Key) = $new
        }
      }
      $matched++
    }
    catch {
      $noResponse.Add([pscustomobject]@{ 物件編號=$record.propertyNo; 案名=$record.caseName; Excel=$candidate.File.FullName; 錯誤=$_.Exception.Message })
    }
    finally {
      if ($null -ne $sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
      if ($null -ne $workbook) { $workbook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    }
  }
}
finally {
  $excel.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

$document.exportedAt = (Get-Date).ToString('o')
$document | ConvertTo-Json -Depth 100 | Set-Content -Encoding UTF8 -LiteralPath $OutputJson
$audit | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $AuditCsv
$unmatched | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath ($AuditCsv -replace '\.csv$', '_找不到Excel.csv')
$noResponse | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath ($AuditCsv -replace '\.csv$', '_無法讀取.csv')

[pscustomobject]@{
  委託中=$active.Count
  已配對Excel=$matched
  修正欄位數=$audit.Count
  找不到Excel=$unmatched.Count
  無法讀取=$noResponse.Count
  修正JSON=$OutputJson
  核對報告=$AuditCsv
} | Format-List



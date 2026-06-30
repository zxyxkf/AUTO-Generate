param(
  [string]$InputJsonPath
)

$ErrorActionPreference = "Stop"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

function Read-JsonInput {
  if (-not [string]::IsNullOrWhiteSpace($InputJsonPath)) {
    $raw = [System.IO.File]::ReadAllText($InputJsonPath, [System.Text.Encoding]::UTF8)
  } else {
    $raw = [Console]::In.ReadToEnd()
  }
  if ([string]::IsNullOrWhiteSpace($raw)) { throw "missing json input" }
  return $raw | ConvertFrom-Json
}

function To-Hash($value) {
  if ($null -eq $value) { return $null }
  if ($value -is [pscustomobject]) {
    $hash = @{}
    foreach ($prop in $value.PSObject.Properties) { $hash[$prop.Name] = To-Hash $prop.Value }
    return $hash
  }
  if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
    $list = @()
    foreach ($item in $value) { $list += To-Hash $item }
    return $list
  }
  return $value
}

function Safe-Text($value) {
  if ($null -eq $value) { return "" }
  return [string]$value
}

function Number-Value($value) {
  if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return 0 }
  $text = ([string]$value) -replace '[^\d\.\-]', ''
  $number = 0.0
  [void][double]::TryParse($text, [ref]$number)
  return $number
}

function Row-HasValue($row) {
  if ($null -eq $row) { return $false }
  foreach ($name in "name","spec","quantity","unitPrice","amount","remark") {
    if ($row.ContainsKey($name) -and -not [string]::IsNullOrWhiteSpace([string]$row[$name])) { return $true }
  }
  return $false
}

function Row-Amount($row, $qty, $price) {
  if ($null -ne $row -and $row.ContainsKey("amount") -and -not [string]::IsNullOrWhiteSpace([string]$row["amount"])) {
    return Number-Value $row.amount
  }
  return $qty * $price
}

function Set-NumberCell($ws, $address, $value, $format) {
  $cell = $ws.Range($address)
  $cell.NumberFormatLocal = $format
  $cell.Value2 = [double](Number-Value $value)
}

function Set-OptionalNumberCell($ws, $address, $value, $format) {
  if ([string]::IsNullOrWhiteSpace([string]$value)) {
    $ws.Range($address).ClearContents() | Out-Null
    return
  }
  Set-NumberCell $ws $address $value $format
}

function Clear-Cells($ws, $addresses) {
  foreach ($address in $addresses) {
    $range = $ws.Range($address)
    if ($range.MergeCells) {
      $range.MergeArea.ClearContents() | Out-Null
    } else {
      $range.ClearContents() | Out-Null
    }
  }
}

function Format-PlainNumber($value) {
  $number = Number-Value $value
  if ([Math]::Abs($number - [Math]::Round($number)) -lt 0.000001) {
    return ([int64][Math]::Round($number)).ToString()
  }
  return $number.ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Format-OptionalPlainNumber($value) {
  if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return "" }
  return Format-PlainNumber $value
}

function Format-MoneyText($value) {
  $number = Number-Value $value
  if ($number -eq 0) { return "" }
  return "￥" + $number.ToString("#,##0.00", [System.Globalization.CultureInfo]::GetCultureInfo("zh-CN"))
}

function ConvertTo-RmbUpper($value) {
  $amount = [int64][Math]::Round((Number-Value $value) * 100)
  if ($amount -eq 0) { return "零元整" }
  $digits = @("零","壹","贰","叁","肆","伍","陆","柒","捌","玖")
  $fractionUnits = @("角","分")
  $integerUnits = @("元","万","亿")
  $smallUnits = @("","拾","佰","仟")
  $result = ""
  for ($i = 0; $i -lt 2; $i++) {
    $number = [int](($amount / [Math]::Pow(10, 1 - $i)) % 10)
    if ($number -ne 0) { $result += $digits[$number] + $fractionUnits[$i] }
  }
  if ([string]::IsNullOrEmpty($result)) { $result = "整" }
  $integer = [int64][Math]::Floor($amount / 100)
  $unitIndex = 0
  while ($integer -gt 0 -and $unitIndex -lt $integerUnits.Count) {
    $part = ""
    for ($j = 0; $j -lt 4 -and $integer -gt 0; $j++) {
      $number = [int]($integer % 10)
      $part = $digits[$number] + $smallUnits[$j] + $part
      $integer = [int64][Math]::Floor($integer / 10)
    }
    $part = [regex]::Replace($part, "(零.)*零$", "")
    if ([string]::IsNullOrEmpty($part)) { $part = "零" }
    $result = $part + $integerUnits[$unitIndex] + $result
    $unitIndex++
  }
  $result = [regex]::Replace($result, "零(拾|佰|仟)", "零")
  $result = [regex]::Replace($result, "零+", "零")
  $result = [regex]::Replace($result, "零(万|亿|元)", '$1')
  $result = $result -replace "亿万", "亿"
  $result = $result -replace "^元", "零元"
  $result = $result -replace "零角零分$", "整"
  $result = $result -replace "零分$", ""
  $result = $result -replace "零角", "零"
  return $result
}

function Full-Path($rootDir, $path) {
  if ([string]::IsNullOrWhiteSpace($path)) { return "" }
  if ([System.IO.Path]::IsPathRooted($path)) { return $path }
  return Join-Path $rootDir $path
}

function New-OutputPaths($rootDir, $type, $payload, $extension) {
  $date = Get-Date -Format "yyyyMMdd"
  $typeLabel = switch ($type) {
    "quotation" { "报价单" }
    "supply" { "供货清单" }
    "contract" { "合同" }
    default { $type }
  }
  $name = if ($payload.customerName) { $payload.customerName } elseif ($payload.buyer) { $payload.buyer } else { "" }
  $name = (Safe-Text $name) -replace '[\\/:*?"<>|]', ''
  if ([string]::IsNullOrWhiteSpace($name)) { $name = "未命名" }
  $dir = Join-Path $rootDir "generated\$date"
  New-Item -ItemType Directory -Force -Path $dir | Out-Null
  $stamp = Get-Date -Format "HHmmss"
  $base = "$typeLabel-$name-$stamp"
  return @{
    Office = Join-Path $dir "$base.$extension"
    Pdf = Join-Path $dir "$base.pdf"
  }
}

function ConvertTo-UrlPath($relativePath) {
  $segments = (Safe-Text $relativePath) -split '/'
  $escaped = @()
  foreach ($segment in $segments) {
    $escaped += [System.Uri]::EscapeDataString($segment)
  }
  return "/" + ($escaped -join "/")
}

function Set-Cell($ws, $address, $value) {
  $ws.Range($address).Value2 = Safe-Text $value
}

function Add-ExcelStamp($ws, $stampPath, $cellAddress, $width, $height) {
  if ([string]::IsNullOrWhiteSpace($stampPath) -or -not (Test-Path -LiteralPath $stampPath)) { return $false }
  $cell = $ws.Range($cellAddress)
  $shape = $ws.Shapes.AddPicture($stampPath, $false, $true, $cell.Left + 16, $cell.Top - 8, $width, $height)
  $shape.Placement = 1
  $shape.ZOrder(0) | Out-Null
  return $true
}

function Add-ExcelStampInRange($ws, $stampPath, $rangeAddress, $width, $height, $offsetRight = 16, $offsetTop = -18) {
  if ([string]::IsNullOrWhiteSpace($stampPath) -or -not (Test-Path -LiteralPath $stampPath)) { return $false }
  $range = $ws.Range($rangeAddress)
  $left = $range.Left + $range.Width - $width - $offsetRight
  $top = $range.Top + $offsetTop
  $shape = $ws.Shapes.AddPicture($stampPath, $false, $true, $left, $top, $width, $height)
  $shape.Placement = 1
  $shape.ZOrder(0) | Out-Null
  return $true
}

function Add-ExcelStampOverRange($ws, $stampPath, $rangeAddress, $width, $height, $offsetRight = 95, $offsetTop = -20) {
  if ([string]::IsNullOrWhiteSpace($stampPath) -or -not (Test-Path -LiteralPath $stampPath)) { return $false }
  $range = $ws.Range($rangeAddress)
  $left = $range.Left + $range.Width - $width - $offsetRight
  $top = $range.Top + $offsetTop
  $shape = $ws.Shapes.AddPicture($stampPath, $false, $true, $left, $top, $width, $height)
  $shape.Placement = 1
  $shape.ZOrder(0) | Out-Null
  return $true
}

function Add-ExcelTextInRange($ws, $text, $rangeAddress, $width, $height, $offsetRight = 110, $offsetTop = 8) {
  $range = $ws.Range($rangeAddress)
  $left = $range.Left + $range.Width - $width - $offsetRight
  $top = $range.Top + $offsetTop
  $shape = $ws.Shapes.AddTextbox(1, $left, $top, $width, $height)
  $shape.Line.Visible = 0
  $shape.Fill.Visible = 0
  $shape.Placement = 1
  $shape.TextFrame.Characters().Text = Safe-Text $text
  $shape.TextFrame.HorizontalAlignment = -4152
  $shape.TextFrame.VerticalAlignment = -4108
  $shape.TextFrame.MarginLeft = 0
  $shape.TextFrame.MarginRight = 0
  $shape.TextFrame.MarginTop = 0
  $shape.TextFrame.MarginBottom = 0
  $shape.TextFrame.Characters().Font.Name = "宋体"
  $shape.TextFrame.Characters().Font.Size = 10
  $shape.TextFrame.Characters().Font.Bold = $true
  return $shape
}

function Set-OnePagePrintArea($ws, $area) {
  $ws.PageSetup.PrintArea = $area
  $ws.PageSetup.Zoom = $false
  $ws.PageSetup.FitToPagesWide = 1
  $ws.PageSetup.FitToPagesTall = 1
  $ws.PageSetup.CenterHorizontally = $true
  $ws.PageSetup.CenterVertically = $false
}

function Set-FixedScalePrintArea($ws, $area) {
  $ws.PageSetup.PrintArea = $area
  $ws.PageSetup.PaperSize = 9
  $ws.PageSetup.Orientation = 1
  $ws.PageSetup.CenterHorizontally = $true
  $ws.PageSetup.CenterVertically = $false
  $ws.PageSetup.FitToPagesWide = $false
  $ws.PageSetup.FitToPagesTall = $false
  $printRange = $ws.Range($area)
  $printableWidth = $ws.Application.InchesToPoints(8.27) - $ws.PageSetup.LeftMargin - $ws.PageSetup.RightMargin
  $zoom = [Math]::Floor(($printableWidth / $printRange.Width) * 100)
  $zoom = [Math]::Max(10, [Math]::Min(100, [int]$zoom))
  $ws.PageSetup.Zoom = $zoom
}

function Trim-WorksheetToPrintRange($ws, $lastColumn, $lastRow) {
  $used = $ws.UsedRange
  $lastUsedRow = $used.Row + $used.Rows.Count - 1
  $lastUsedColumn = $used.Column + $used.Columns.Count - 1
  if ($lastUsedColumn -gt $lastColumn) {
    $startColumn = $lastColumn + 1
    [void]$ws.Range($ws.Cells.Item(1, $startColumn), $ws.Cells.Item(1, $lastUsedColumn)).EntireColumn.Delete()
  }
  if ($lastUsedRow -gt $lastRow) {
    $startRow = $lastRow + 1
    [void]$ws.Rows("${startRow}:${lastUsedRow}").Delete()
  }
  [void]$ws.UsedRange
}

function Set-DetailRowCount($ws, $startRow, $templateRowCount, $desiredRowCount, $copySourceRow = 0) {
  $count = [Math]::Max(1, [int]$desiredRowCount)
  if ($count -gt $templateRowCount) {
    for ($i = $templateRowCount; $i -lt $count; $i++) {
      $sourceRow = if ($copySourceRow -gt 0) { $copySourceRow } else { $startRow + $templateRowCount - 1 }
      $insertRow = $startRow + $i
      [void]$ws.Rows($sourceRow).Copy()
      [void]$ws.Rows($insertRow).Insert()
      $ws.Application.CutCopyMode = $false
    }
  } elseif ($count -lt $templateRowCount) {
    $firstDelete = $startRow + $count
    $lastDelete = $startRow + $templateRowCount - 1
    [void]$ws.Rows("${firstDelete}:${lastDelete}").Delete()
  }
  return $count
}

function Normalize-QuotationDetailRows($ws, $startRow, $rowCount) {
  for ($i = 0; $i -lt $rowCount; $i++) {
    $r = $startRow + $i
    $range = $ws.Range("B${r}:C${r}")
    if ($range.MergeCells) { [void]$range.UnMerge() }
  }
}

function Reset-SupplyRemarkRange($ws, $startRow, $rowCount, $totalRow) {
  $remarkRange = $ws.Range("G${startRow}:H${totalRow}")
  if ($remarkRange.MergeCells) { [void]$remarkRange.UnMerge() }
  [void]$remarkRange.Merge()
  $remarkRange.HorizontalAlignment = -4108
  $remarkRange.VerticalAlignment = -4108
  $remarkRange.WrapText = $true
}

function Reset-SupplyFooterRow($ws, $row) {
  $rowRange = $ws.Range("A${row}:H${row}")
  [void]$rowRange.UnMerge()
  $leftRange = $ws.Range("A${row}:E${row}")
  [void]$leftRange.Merge()
  $leftRange.ClearContents() | Out-Null
  $rightRange = $ws.Range("F${row}:H${row}")
  [void]$rightRange.Merge()
  Write-Output -NoEnumerate $rightRange
}

function Fill-Quotation($ws, $payload) {
  Set-Cell $ws "A2" "客户名称：$(Safe-Text $payload.customerName)"
  $rows = @($payload.rows)
  $rowCount = Set-DetailRowCount $ws 4 5 $rows.Count 7
  Normalize-QuotationDetailRows $ws 4 $rowCount
  $total = 0
  for ($i = 0; $i -lt $rowCount; $i++) {
    $r = 4 + $i
    $row = if ($i -lt $rows.Count) { $rows[$i] } else { $null }
    Set-Cell $ws "A$r" ($i + 1)
    if (-not (Row-HasValue $row)) {
      Clear-Cells $ws @("B$r","C$r","D$r","E$r","F$r")
      continue
    }
    $qty = Number-Value $row.quantity
    $price = Number-Value $row.unitPrice
    $amount = Row-Amount $row $qty $price
    $total += $amount
    Set-Cell $ws "B$r" (Safe-Text $payload.date)
    Set-Cell $ws "C$r" (Safe-Text $row.name)
    Set-OptionalNumberCell $ws "D$r" $row.unitPrice "0.00"
    Set-OptionalNumberCell $ws "E$r" $row.quantity "0"
    Set-NumberCell $ws "F$r" $amount "￥#,##0.00"
  }
  $totalRow = 4 + $rowCount
  $upperRow = $totalRow + 1
  $supplierRow = $totalRow + 2
  Set-NumberCell $ws "F$totalRow" $total "￥#,##0.00"
  Set-Cell $ws "A$upperRow" "大写金额：$(ConvertTo-RmbUpper $total)"
  Clear-Cells $ws @("A$supplierRow")
  $supplierText = "供货方（盖章）：$(Safe-Text $payload.company)"
  [void](Add-ExcelTextInRange $ws $supplierText "A${supplierRow}:F${supplierRow}" 360 18 34 13)
  return @{
    PrintArea = "A1:F$($supplierRow + 1)"
    StampRange = "A${supplierRow}:F${supplierRow}"
  }
}

function Fill-Supply($ws, $payload) {
  Set-Cell $ws "A2" "供货时间：$(Safe-Text $payload.date)"
  $rows = @($payload.rows)
  $rowCount = Set-DetailRowCount $ws 4 3 $rows.Count
  $total = 0
  for ($i = 0; $i -lt $rowCount; $i++) {
    $r = 4 + $i
    $row = if ($i -lt $rows.Count) { $rows[$i] } else { $null }
    Set-Cell $ws "A$r" ($i + 1)
    if (-not (Row-HasValue $row)) {
      Clear-Cells $ws @("B$r","C$r","D$r","E$r","F$r")
      continue
    }
    $qty = Number-Value $row.quantity
    $price = Number-Value $row.unitPrice
    $amount = Row-Amount $row $qty $price
    $total += $amount
    Set-Cell $ws "B$r" (Safe-Text $row.name)
    Set-Cell $ws "C$r" (Safe-Text $row.spec)
    Set-OptionalNumberCell $ws "D$r" $row.quantity "0"
    Set-OptionalNumberCell $ws "E$r" $row.unitPrice "0.00"
    Set-NumberCell $ws "F$r" $amount "￥#,##0.00"
  }
  $totalRow = 4 + $rowCount
  Reset-SupplyRemarkRange $ws 4 $rowCount $totalRow
  Set-Cell $ws "G4" (Safe-Text $payload.remark)
  Set-Cell $ws "A$totalRow" "合计实收：$(Format-MoneyText $total)"
  Set-Cell $ws "C$totalRow" "大写：$(ConvertTo-RmbUpper $total)"
  $infoRow = $totalRow + 1
  $stampRow = $totalRow + 2
  $infoRange = Reset-SupplyFooterRow $ws $infoRow
  $infoRange.HorizontalAlignment = -4131
  $infoRange.VerticalAlignment = -4160
  $infoRange.WrapText = $true
  Set-Cell $ws "F$infoRow" "供货单位：`n$(Safe-Text $payload.supplierInfo)"
  $stampRange = Reset-SupplyFooterRow $ws $stampRow
  $stampRange.HorizontalAlignment = -4131
  $stampRange.VerticalAlignment = -4108
  Set-Cell $ws "F$stampRow" "供货单位（盖章）："
  $printBottom = $stampRow + 3
  return @{
    PrintArea = "A1:H$printBottom"
    StampRange = "F${stampRow}:H${stampRow}"
    PrintBottom = $printBottom
  }
}

function Generate-Excel($rootDir, $type, $payload) {
  $template = Full-Path $rootDir $payload.templatePath
  if (-not (Test-Path -LiteralPath $template)) { throw "模板不存在：$template" }
  $paths = New-OutputPaths $rootDir $type $payload "xlsx"
  Copy-Item -LiteralPath $template -Destination $paths.Office -Force

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $null
  try {
    $wb = $excel.Workbooks.Open($paths.Office)
    $ws = $wb.Worksheets.Item(1)
    if ($type -eq "quotation") {
      $layout = Fill-Quotation $ws $payload
      Set-OnePagePrintArea $ws $layout.PrintArea
      [void](Add-ExcelStampInRange $ws $payload.stampPath $layout.StampRange 76 76 34 -20)
    } else {
      $layout = Fill-Supply $ws $payload
      Trim-WorksheetToPrintRange $ws 8 $layout.PrintBottom
      Set-FixedScalePrintArea $ws $layout.PrintArea
      [void](Add-ExcelStampOverRange $ws $payload.stampPath $layout.StampRange 82 82 56 -22)
    }
    $wb.SaveAs($paths.Office, 51)
    $ws.ExportAsFixedFormat(0, $paths.Pdf)
    $wb.Close($true)
  } finally {
    if ($wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    try { $excel.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  return @{ officePath = $paths.Office; pdfPath = $paths.Pdf; stampUsed = [bool]$payload.stampPath }
}

function Replace-All($doc, $find, $replace) {
  $range = $doc.Content
  $finder = $range.Find
  $finder.ClearFormatting()
  $finder.Replacement.ClearFormatting()
  [void]$finder.Execute($find, $false, $false, $false, $false, $false, $true, 1, $false, $replace, 2)
}

function Normalize-WordText($value) {
  return (Safe-Text $value) -replace "(\r\n|\n|\r)", "`r"
}

function Set-WordCellText($cell, $text) {
  $range = $cell.Range.Duplicate
  if ($range.End -gt $range.Start) { $range.End = $range.End - 1 }
  $range.Text = Normalize-WordText $text
}

function Set-ContractTopField($doc, $label, $value) {
  $text = Safe-Text $value
  $content = $doc.Content
  $finder = $content.Find
  $finder.ClearFormatting()
  [void]$finder.Execute($label)
  if (-not $finder.Found) { return }
  $paragraph = $content.Paragraphs.Item(1).Range
  $start = $content.End
  $end = $paragraph.End - 1
  if ($end -le $start) { return }
  $fieldRange = $doc.Range($start, $end)
  $fieldRange.Text = $text
  $fillRange = $doc.Range($start, $start + $text.Length)
  $fillRange.Font.Underline = 0
}

function Clear-ContractTopUnderlines($doc) {
  $firstTableStart = if ($doc.Tables.Count -ge 1) { $doc.Tables.Item(1).Range.Start } else { $doc.Content.End }
  foreach ($paragraph in $doc.Paragraphs) {
    $range = $paragraph.Range
    if ($range.Start -ge $firstTableStart) { break }
    $text = Safe-Text $range.Text
    if ($text.Contains("需方：") -or $text.Contains("供方：")) {
      $cleanRange = $range.Duplicate
      if ($cleanRange.End -gt $cleanRange.Start) { $cleanRange.End = $cleanRange.End - 1 }
      $cleanRange.Font.Underline = 0
    }
  }
}

function Add-WordStamp($doc, $stampPath) {
  if ([string]::IsNullOrWhiteSpace($stampPath) -or -not (Test-Path -LiteralPath $stampPath)) { return $false }
  $width = 76
  $height = 76
  $left = 360
  $top = 590
  if ($doc.Tables.Count -ge 2) {
    try {
      $cell = $doc.Tables.Item(2).Cell(1, 2)
      $anchor = $cell.Range.Paragraphs.Item(1).Range
      $cellLeft = [double]$anchor.Information(5)
      $cellTop = [double]$anchor.Information(6)
      if ($cellLeft -gt 0 -and $cellTop -gt 0) {
        $left = $cellLeft + (([double]$cell.Width - $width) / 2)
        $top = $cellTop + (([double]$cell.Height - $height) / 2)
      }
    } catch {}
  }
  $shape = $doc.Shapes.AddPicture($stampPath, $false, $true, $left, $top, $width, $height)
  $shape.WrapFormat.Type = 3
  try { $shape.ZOrder(0) | Out-Null } catch {}
  return $true
}

function Fill-ContractSupplier($doc, $company) {
  $text = Safe-Text $company
  if ([string]::IsNullOrWhiteSpace($text)) { return }
  $knownSuppliers = @(
    "Example Supplier Ltd.",
    "Example VAT Supplier Ltd."
  )
  foreach ($supplier in $knownSuppliers) {
    if ($supplier -ne $text) { Replace-All $doc $supplier $text }
  }
  Replace-All $doc "供方：                    签订时间" "供方：$text                    签订时间"
  Replace-All $doc "供方：                                 签订时间" "供方：$text                    签订时间"
}

function Remove-ContractLocation($doc) {
  Replace-All $doc "签订地点：" ""
}

function Fill-ContractTable($doc, $payload) {
  if ($doc.Tables.Count -lt 1) { return 0 }
  $table = $doc.Tables.Item(1)
  $items = @()
  if ($payload.ContainsKey("items") -and $null -ne $payload.items) {
    foreach ($item in @($payload.items)) {
      if ($null -ne $item) { $items += $item }
    }
  }
  $desiredRows = [Math]::Max(1, $items.Count)
  $currentDetailRows = [Math]::Max(1, $table.Rows.Count - 2)
  if ($desiredRows -gt $currentDetailRows) {
    for ($i = $currentDetailRows; $i -lt $desiredRows; $i++) {
      $insertAfterRow = 1 + $i
      $table.Cell($insertAfterRow, 1).Select()
      $doc.Application.Selection.InsertRowsBelow(1) | Out-Null
    }
  }
  $detailRows = [Math]::Max(1, $table.Rows.Count - 2)
  $totalRow = $table.Rows.Count
  $total = 0
  for ($i = 0; $i -lt $detailRows; $i++) {
    $rowIndex = 2 + $i
    $item = if ($i -lt $items.Count) { $items[$i] } else { $null }
    if (-not (Row-HasValue $item)) {
      Set-WordCellText $table.Cell($rowIndex, 1) ""
      Set-WordCellText $table.Cell($rowIndex, 2) ""
      Set-WordCellText $table.Cell($rowIndex, 3) ""
      Set-WordCellText $table.Cell($rowIndex, 4) ""
      continue
    }
    $qty = Number-Value $item.quantity
    $price = Number-Value $item.unitPrice
    $amount = Row-Amount $item $qty $price
    $total += $amount
    Set-WordCellText $table.Cell($rowIndex, 1) (Safe-Text $item.name)
    Set-WordCellText $table.Cell($rowIndex, 2) (Format-OptionalPlainNumber $item.quantity)
    Set-WordCellText $table.Cell($rowIndex, 3) (Format-OptionalPlainNumber $item.unitPrice)
    Set-WordCellText $table.Cell($rowIndex, 4) (Format-MoneyText $amount)
  }
  try { Set-WordCellText $table.Cell(2, 5) (Safe-Text $payload.remark) } catch {}
  if ($totalRow -ge 3) {
    Set-WordCellText $table.Cell($totalRow, 1) "大写金额：$(ConvertTo-RmbUpper $total)"
    Set-WordCellText $table.Cell($totalRow, 2) "人民币：$(Format-MoneyText $total) 元"
  }
  return $total
}

function Fill-ContractSignArea($doc, $payload) {
  if ($doc.Tables.Count -lt 2) { return }
  $table = $doc.Tables.Item(2)
  $buyerText = "需方"
  if (-not [string]::IsNullOrWhiteSpace([string]$payload.buyerInfo)) {
    $buyerText = "$buyerText`r$(Safe-Text $payload.buyerInfo)"
  }
  $supplierText = "供方"
  if (-not [string]::IsNullOrWhiteSpace([string]$payload.supplierInfo)) {
    $supplierText = "$supplierText`r$(Safe-Text $payload.supplierInfo)"
  }
  Set-WordCellText $table.Cell(1, 1) $buyerText
  Set-WordCellText $table.Cell(1, 2) $supplierText
}

function Get-ContractPaymentMethod($payload) {
  $offlineIds = @("example_invoice", "example_vat")
  $offlineLabels = @("Example Invoice", "Example VAT")
  $themeId = Safe-Text $payload.themeId
  $themeLabel = Safe-Text $payload.themeLabel
  if ($offlineIds -contains $themeId -or $offlineLabels -contains $themeLabel) { return "对公线下付款" }
  return "淘宝平台下单"
}

function Get-ContractPackaging($payload) {
  $text = Safe-Text $payload.packaging
  if ([string]::IsNullOrWhiteSpace($text)) { return "塑料袋包装" }
  return $text.Trim()
}

function Get-ContractClausesText($payload) {
  $payment = Get-ContractPaymentMethod $payload
  $packaging = Get-ContractPackaging $payload
  return @(
    "二、结算方式与发货时间：",
    "    $payment，定制产品3-5天左右出货。",
    "三、包装、运输方式：",
    "    $packaging，圆通快递包邮。",
    "四、违约责任：",
    "双方应严格遵守本合同的约定，如出现问题，根据问题所属承担责任。解决合同纠纷的",
    "方式：协商调解不成时，依法向双方所在地人民法院诉讼。",
    "五、其它事项：",
    "1.定制产品不支持退货，保证质量。",
    "2.如货物出现质量问题，需方在收到货七天内通知供方，供方负责调换并承担相关运输费用。",
    "3.供方保证产品符合相关国家和行业标准，因供方产品给需方造成的损失，供方承担赔偿责任及包括律师费在内的维权成本。",
    "六、本合同邮寄具有法律效力，需方收到邮寄签字盖章回传后生效。"
  ) -join "`r"
}

function Set-ContractClauseParagraph($doc, $keyword, $text, $manualPrefix = "") {
  if ($doc.Tables.Count -lt 2) { return $false }
  foreach ($paragraph in $doc.Paragraphs) {
    $range = $paragraph.Range
    if ($range.Start -lt $doc.Tables.Item(1).Range.End) { continue }
    if ($range.Start -ge $doc.Tables.Item(2).Range.Start) { break }
    if ((Safe-Text $range.Text).Contains($keyword)) {
      $hasListNumber = $false
      try { $hasListNumber = -not [string]::IsNullOrWhiteSpace([string]$range.ListFormat.ListString) } catch {}
      $finalText = if ($hasListNumber -or [string]::IsNullOrWhiteSpace($manualPrefix)) { $text } else { "$manualPrefix$text" }
      $start = $range.Start
      $range.Text = "$(Normalize-WordText $finalText)`r"
      $newRange = $doc.Range($start, [Math]::Min($start + $finalText.Length + 1, $doc.Content.End))
      if (-not $hasListNumber) {
        try { $newRange.ListFormat.RemoveNumbers() | Out-Null } catch {}
      }
      foreach ($newParagraph in $newRange.Paragraphs) {
        if (-not $hasListNumber) {
          $newParagraph.LeftIndent = 0
          $newParagraph.FirstLineIndent = 0
        }
        $newParagraph.Alignment = 0
        $newParagraph.SpaceBefore = 0
        $newParagraph.SpaceAfter = 0
      }
      return $true
    }
  }
  return $false
}

function Fill-ContractClauses($doc, $payload) {
  if ($doc.Tables.Count -lt 2) { return }
  $lines = @(
    "二、结算方式与发货时间：",
    "    $(Get-ContractPaymentMethod $payload)，定制产品3-5天左右出货。",
    "三、包装、运输方式：",
    "    $(Get-ContractPackaging $payload)，圆通快递包邮。",
    "四、违约责任：",
    "双方应严格遵守本合同的约定，如出现问题，根据问题所属承担责任。解决合同纠纷的",
    "方式：协商调解不成时，依法向双方所在地人民法院诉讼。",
    "五、其它事项：",
    "1.定制产品不支持退货，保证质量。",
    "2.如货物出现质量问题，需方在收到货七天内通知供方，供方负责调换并承担相关运输费用。",
    "3.供方保证产品符合相关国家和行业标准，因供方产品给需方造成的损失，供方承担赔偿责任及包括律师费在内的维权成本。",
    "六、本合同邮寄具有法律效力，需方收到邮寄签字盖章回传后生效。"
  )
  $paragraphs = @()
  $start = $doc.Tables.Item(1).Range.End
  $end = $doc.Tables.Item(2).Range.Start
  foreach ($paragraph in $doc.Paragraphs) {
    if ($paragraph.Range.Start -lt $start) { continue }
    if ($paragraph.Range.Start -ge $end) { break }
    $paragraphs += $paragraph
  }
  $count = [Math]::Min($paragraphs.Count, $lines.Count)
  if ($count -eq 0) { return }
  for ($i = $count - 1; $i -ge 0; $i--) {
    $range = $paragraphs[$i].Range
    try { $range.ListFormat.RemoveNumbers() | Out-Null } catch {}
    if ($i -eq $count - 1 -and $count -lt $lines.Count) {
      $replacement = (($lines[$i..($lines.Count - 1)] | ForEach-Object { Normalize-WordText $_ }) -join "`r") + "`r"
    } else {
      $replacement = "$(Normalize-WordText $lines[$i])`r"
    }
    $range.Text = $replacement
  }

  $extraParagraphs = @()
  $updatedStart = $doc.Tables.Item(1).Range.End
  $updatedEnd = $doc.Tables.Item(2).Range.Start
  foreach ($paragraph in $doc.Paragraphs) {
    if ($paragraph.Range.Start -lt $updatedStart) { continue }
    if ($paragraph.Range.Start -ge $updatedEnd) { break }
    $cleanText = (Safe-Text $paragraph.Range.Text).Trim()
    if ([string]::IsNullOrWhiteSpace($cleanText)) { continue }
    $extraParagraphs += $paragraph
    try { $paragraph.Range.ListFormat.RemoveNumbers() | Out-Null } catch {}
    $paragraph.LeftIndent = 0
    $paragraph.FirstLineIndent = 0
    $paragraph.Alignment = 0
    $paragraph.SpaceBefore = 0
    $paragraph.SpaceAfter = 0
  }
  for ($i = $extraParagraphs.Count - 1; $i -ge $lines.Count; $i--) {
    $range = $extraParagraphs[$i].Range
    try { $range.ListFormat.RemoveNumbers() | Out-Null } catch {}
    [void]$range.Delete()
  }
}

function Generate-Word($rootDir, $payload) {
  $template = Full-Path $rootDir $payload.templatePath
  if (-not (Test-Path -LiteralPath $template)) { throw "模板不存在：$template" }
  $paths = New-OutputPaths $rootDir "contract" $payload "docx"

  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $null
  try {
    $doc = $word.Documents.Open($template)
    $contractTotal = 0
    if ($payload.items) {
      foreach ($item in @($payload.items)) {
        $contractTotal += (Number-Value $item.quantity) * (Number-Value $item.unitPrice)
      }
    }
    Set-ContractTopField $doc "需方：" $payload.buyer
    Fill-ContractSupplier $doc $payload.company
    Remove-ContractLocation $doc
    Replace-All $doc "签订时间：年月日" "签订时间：$(Safe-Text $payload.date)"
    Clear-ContractTopUnderlines $doc
    [void](Fill-ContractTable $doc $payload)
    Fill-ContractClauses $doc $payload
    Fill-ContractSignArea $doc $payload
    [void](Add-WordStamp $doc $payload.stampPath)
    $doc.SaveAs([ref]$paths.Office, [ref]16)
    $doc.ExportAsFixedFormat($paths.Pdf, 17)
    $doc.Close($true)
  } finally {
    if ($doc) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
    try { $word.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  }
  return @{ officePath = $paths.Office; pdfPath = $paths.Pdf; stampUsed = [bool]$payload.stampPath }
}

$inputData = Read-JsonInput
$rootDir = [string]$inputData.rootDir
$type = [string]$inputData.type
$payload = To-Hash $inputData.payload

if ($type -eq "contract") {
  $result = Generate-Word $rootDir $payload
} else {
  $result = Generate-Excel $rootDir $type $payload
}

$result.relativeOfficePath = $result.officePath.Substring($rootDir.Length).TrimStart('\') -replace '\\','/'
$result.relativePdfPath = $result.pdfPath.Substring($rootDir.Length).TrimStart('\') -replace '\\','/'
$result.officeUrl = ConvertTo-UrlPath $result.relativeOfficePath
$result.pdfUrl = ConvertTo-UrlPath $result.relativePdfPath
$result | ConvertTo-Json -Depth 6 -Compress


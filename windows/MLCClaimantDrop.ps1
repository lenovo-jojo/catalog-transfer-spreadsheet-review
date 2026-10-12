param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths
)

Add-Type -AssemblyName System.Windows.Forms

function Show-Message([string]$Text, [string]$Title, [System.Windows.Forms.MessageBoxIcon]$Icon) {
    [void][System.Windows.Forms.MessageBox]::Show(
        $Text,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $Icon
    )
}

function Csv-Escape([string]$Value) {
    if ($null -eq $Value) { return '' }
    if ($Value.Contains(',') -or $Value.Contains('"') -or $Value.Contains("`r") -or $Value.Contains("`n")) {
        return '"' + $Value.Replace('"', '""') + '"'
    }
    return $Value
}

if (-not $Paths -or $Paths.Count -eq 0) {
    Show-Message 'Drag one MLC .xlsx report onto the desktop icon.' 'MLC Claimant Drop' ([System.Windows.Forms.MessageBoxIcon]::Information)
    exit 0
}

$inputPath = [System.IO.Path]::GetFullPath($Paths[0])
if ([System.IO.Path]::GetExtension($inputPath).ToLowerInvariant() -ne '.xlsx') {
    Show-Message 'Please drop an .xlsx report onto the icon.' 'MLC Claimant Drop' ([System.Windows.Forms.MessageBoxIcon]::Warning)
    exit 1
}

$excel = $null
$workbook = $null
$sheet = $null
$range = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($inputPath, 0, $true)
    $sheet = $workbook.Worksheets.Item(1)
    $lastRow = $sheet.UsedRange.Row + $sheet.UsedRange.Rows.Count - 1

    if ($lastRow -lt 2) {
        throw 'No report rows were found.'
    }

    $range = $sheet.Range("A1:R$lastRow")
    $data = $range.Value2
    $columns = @(11, 12, 13, 14, 17, 18)
    $fallbackHeaders = @(
        'MEMBER MLC P#',
        'MEMBER NAME',
        'PUBLISHER MLC P#',
        'PUBLISHER NAME',
        'ADMINISTRATOR MLC P#',
        'ADMINISTRATOR NAME'
    )

    $headers = @()
    for ($columnIndex = 0; $columnIndex -lt $columns.Count; $columnIndex++) {
        $header = [string]$data[1, $columns[$columnIndex]]
        if ([string]::IsNullOrWhiteSpace($header)) { $header = $fallbackHeaders[$columnIndex] }
        $headers += $header
    }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $uniqueRows = New-Object 'System.Collections.Generic.List[object]'
    $matchedRows = 0

    for ($rowNumber = 2; $rowNumber -le $lastRow; $rowNumber++) {
        $answer = ([string]$data[$rowNumber, 5]).Trim().ToUpperInvariant()
        $answer = $answer.Replace('(', '').Replace(')', '').Trim()
        if ($answer -ne 'YES' -and $answer -ne 'Y') { continue }

        $matchedRows++
        $claimant = @()
        foreach ($column in $columns) {
            $claimant += [string]$data[$rowNumber, $column]
        }

        $keyText = $claimant -join [char]31
        $key = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($keyText))
        if ($seen.Add($key)) {
            [void]$uniqueRows.Add($claimant)
        }
    }

    if ($matchedRows -eq 0) {
        throw 'No YES, Yes, Y, or (Y) rows were found in column E.'
    }

    $lines = New-Object 'System.Collections.Generic.List[string]'
    [void]$lines.Add(($headers -join "`t"))
    foreach ($claimant in $uniqueRows) {
        [void]$lines.Add(($claimant -join "`t"))
    }
    $clipboardText = $lines -join "`r`n"
    Set-Clipboard -Value $clipboardText

    $csvLines = New-Object 'System.Collections.Generic.List[string]'
    [void]$csvLines.Add((($headers | ForEach-Object { Csv-Escape $_ }) -join ','))
    foreach ($claimant in $uniqueRows) {
        [void]$csvLines.Add((($claimant | ForEach-Object { Csv-Escape ([string]$_) }) -join ','))
    }

    $outputName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath) + '_yes_claimants.csv'
    $outputPath = Join-Path ([System.IO.Path]::GetDirectoryName($inputPath)) $outputName
    [System.IO.File]::WriteAllLines($outputPath, $csvLines, (New-Object System.Text.UTF8Encoding($true)))

    $plural = if ($uniqueRows.Count -eq 1) { 'row' } else { 'rows' }
    Show-Message ("Found $matchedRows matching rows and $($uniqueRows.Count) unique claimant $plural.`r`n`r`nThe table is copied—press Ctrl+V to paste it.`r`n`r`nAlso saved:`r`n$outputPath") 'MLC Claimant Drop' ([System.Windows.Forms.MessageBoxIcon]::Information)
}
catch {
    Show-Message ("The report could not be processed:`r`n`r`n$($_.Exception.Message)") 'MLC Claimant Drop' ([System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}
finally {
    if ($workbook -ne $null) { $workbook.Close($false) }
    if ($excel -ne $null) { $excel.Quit() }

    foreach ($comObject in @($range, $sheet, $workbook, $excel)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

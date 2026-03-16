#requires -Version 5.1

param(
    [Parameter(Mandatory = $true)]
    [string]$OutputXlsx
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-ModuleInstalled {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    return [bool](Get-Module -ListAvailable -Name $Name)
}

function Convert-CsvToXlsx {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $true)]
        [string]$XlsxPath
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = 'GroupMembers'

        $queryTable = $worksheet.QueryTables.Add("TEXT;$CsvPath", $worksheet.Range("A1"))
        $queryTable.TextFileParseType = 1
        $queryTable.TextFileCommaDelimiter = $true
        $queryTable.TextFilePlatform = 65001
        $queryTable.AdjustColumnWidth = $true
        $queryTable.Refresh($false)
        $queryTable.Delete()

        $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        $workbook.SaveAs($XlsxPath, 51)
    }
    finally {
        if ($worksheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) }
        if ($workbook) {
            $workbook.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
        if ($excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if (-not (Test-ModuleInstalled -Name 'AzureAD')) {
    throw "AzureAD module is not installed."
}

Import-Module AzureAD -ErrorAction Stop

# Results folder in current location
$baseFolder = (Get-Location).Path
$resultsFolder = Join-Path -Path $baseFolder -ChildPath "results"
New-Item -ItemType Directory -Path $resultsFolder -Force | Out-Null

# Force output into .\results
$outputFileName = Split-Path -Path $OutputXlsx -Leaf
if ([string]::IsNullOrWhiteSpace($outputFileName)) {
    throw "OutputXlsx must contain a valid file name."
}
$OutputXlsx = Join-Path -Path $resultsFolder -ChildPath $outputFileName

# Safe temp folder
$tempBase = $env:TEMP
if ([string]::IsNullOrWhiteSpace($tempBase)) {
    $tempBase = Join-Path -Path $resultsFolder -ChildPath "temp"
    New-Item -ItemType Directory -Path $tempBase -Force | Out-Null
}

$tempRoot = Join-Path -Path $tempBase -ChildPath ("AzureAD_GroupExport_" + [guid]::NewGuid().Guid)
$tempCsv  = Join-Path -Path $tempRoot -ChildPath "GroupMembers.csv"

New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

# Quick check that a session already exists
try {
    $null = Get-AzureADTenantDetail
}
catch {
    throw "No active AzureAD session detected. Run Connect-AzureAD first in this Cloud Shell session."
}

Write-Host "Retrieving all security-enabled groups..." -ForegroundColor Cyan
$groups = Get-AzureADMSGroup -All $true | Where-Object { $_.SecurityEnabled -eq $true }

$groupCount = @($groups).Count
Write-Host "Found $groupCount security-enabled groups." -ForegroundColor Green

# Write CSV header
"GroupId,MemberId" | Set-Content -LiteralPath $tempCsv -Encoding UTF8

$lineBuffer = New-Object System.Collections.Generic.List[string]
$flushThreshold = 5000
$processed = 0
$failed = New-Object System.Collections.Generic.List[object]

foreach ($group in $groups) {
    try {
        $members = Get-AzureADGroupMember -ObjectId $group.Id -All $true

        foreach ($member in $members) {
            $lineBuffer.Add("$($group.Id),$($member.ObjectId)")
        }

        if ($lineBuffer.Count -ge $flushThreshold) {
            Add-Content -LiteralPath $tempCsv -Value $lineBuffer -Encoding UTF8
            $lineBuffer.Clear()
        }
    }
    catch {
        $failed.Add([pscustomobject]@{
            GroupId = $group.Id
            Error   = $_.Exception.Message
        }) | Out-Null
    }

    $processed++
    if (($processed % 100) -eq 0 -or $processed -eq $groupCount) {
        Write-Host ("Completed {0}/{1} groups..." -f $processed, $groupCount) -ForegroundColor DarkCyan
    }
}

if ($lineBuffer.Count -gt 0) {
    Add-Content -LiteralPath $tempCsv -Value $lineBuffer -Encoding UTF8
    $lineBuffer.Clear()
}

try {
    Write-Host "Converting CSV to XLSX..." -ForegroundColor Cyan
    Convert-CsvToXlsx -CsvPath $tempCsv -XlsxPath $OutputXlsx
    Write-Host "Export complete: $OutputXlsx" -ForegroundColor Green
}
catch {
    $csvFallback = [System.IO.Path]::ChangeExtension($OutputXlsx, '.csv')
    Copy-Item -LiteralPath $tempCsv -Destination $csvFallback -Force
    Write-Warning "Could not create XLSX through Excel COM. CSV created instead: $csvFallback"
    Write-Warning $_.Exception.Message
}

if ($failed.Count -gt 0) {
    $errorLog = Join-Path -Path $resultsFolder -ChildPath 'FailedGroups.csv'
    $failed | Export-Csv -LiteralPath $errorLog -NoTypeInformation -Encoding UTF8
    Write-Warning "$($failed.Count) groups failed. Failure log: $errorLog"
}
else {
    Write-Host "All groups processed successfully." -ForegroundColor Green
}

#requires -Version 5.1

<#
.SYNOPSIS
    Export all security-enabled Azure AD group members to an Excel file (GroupId, MemberId),
    using ONLY the AzureAD module and runspace-based parallelism.

.NOTES
    - Optimized for speed with runspace pool parallelism.
    - Uses AzureAD module only.
    - Writes a CSV first, then converts to XLSX via Excel COM if Excel is installed.
    - Best used with a non-MFA account because each parallel worker reconnects to AzureAD.

.PARAMETER TenantId
    Azure AD tenant ID (GUID).

.PARAMETER OutputXlsx
    Final XLSX output path.

.PARAMETER ThreadCount
    Number of parallel runspaces. Start with 8, 12, or 16 depending on tenant size and throttling.

.EXAMPLE
    .\Export-AzureADSecurityGroupMembers.ps1 `
        -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -OutputXlsx "C:\Temp\AzureAD_SecurityGroupMembers.xlsx" `
        -ThreadCount 12
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$OutputXlsx,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 64)]
    [int]$ThreadCount = 12
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------
# Helpers
# ---------------------------

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

        # Import CSV into the sheet
        $queryTable = $worksheet.QueryTables.Add("TEXT;$CsvPath", $worksheet.Range("A1"))
        $queryTable.TextFileParseType = 1        # xlDelimited
        $queryTable.TextFileCommaDelimiter = $true
        $queryTable.TextFilePlatform = 65001     # UTF-8
        $queryTable.AdjustColumnWidth = $true
        $queryTable.Refresh($false)
        $queryTable.Delete()

        # Optional formatting
        $usedRange = $worksheet.UsedRange
        $usedRange.EntireColumn.AutoFit() | Out-Null

        # 51 = xlOpenXMLWorkbook (.xlsx)
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

# ---------------------------
# Validation
# ---------------------------

if (-not (Test-ModuleInstalled -Name 'AzureAD')) {
    throw "AzureAD module is not installed."
}

Import-Module AzureAD -ErrorAction Stop

# Make sure output folder exists
$outputFolder = Split-Path -Path $OutputXlsx -Parent
if (-not [string]::IsNullOrWhiteSpace($outputFolder) -and -not (Test-Path -LiteralPath $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
}

$tempRoot = Join-Path $env:TEMP ("AzureAD_GroupExport_" + [guid]::NewGuid().Guid)
$tempCsv  = Join-Path $tempRoot "GroupMembers.csv"
$tempDir  = Join-Path $tempRoot "Chunks"

New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# ---------------------------
# Connect once in the main session
# ---------------------------

Write-Host "Prompting for Azure AD credential..." -ForegroundColor Cyan
$credential = Get-Credential -Message "Enter Azure AD credential for tenant $TenantId"

Write-Host "Connecting to Azure AD in main session..." -ForegroundColor Cyan
Connect-AzureAD -TenantId $TenantId -Credential $credential | Out-Null

# ---------------------------
# Get all security-enabled groups
# ---------------------------
# AzureAD/MS cmdlets expose SecurityEnabled on group objects. We use Get-AzureADMSGroup
# because it returns modern group properties reliably within the AzureAD module family.
# Filtering client-side avoids relying on endpoint/filter quirks across tenants.

Write-Host "Retrieving all security-enabled groups..." -ForegroundColor Cyan

$groups = Get-AzureADMSGroup -All $true | Where-Object { $_.SecurityEnabled -eq $true }

$groupCount = @($groups).Count
Write-Host "Found $groupCount security-enabled groups." -ForegroundColor Green

# Create CSV header up front
"GroupId,MemberId" | Set-Content -LiteralPath $tempCsv -Encoding UTF8

if ($groupCount -eq 0) {
    Write-Host "No security-enabled groups found. Creating empty output." -ForegroundColor Yellow
    try {
        Convert-CsvToXlsx -CsvPath $tempCsv -XlsxPath $OutputXlsx
        Write-Host "Done: $OutputXlsx" -ForegroundColor Green
    }
    catch {
        Copy-Item -LiteralPath $tempCsv -Destination ([System.IO.Path]::ChangeExtension($OutputXlsx, '.csv')) -Force
        Write-Warning "Excel COM not available. CSV created instead: $([System.IO.Path]::ChangeExtension($OutputXlsx, '.csv'))"
    }
    return
}

# ---------------------------
# Runspace worker
# Each worker:
#   - Imports AzureAD
#   - Connects with supplied credential
#   - Gets all members for one group
#   - Writes one chunk CSV file
# ---------------------------

$workerScript = {
    param(
        [string]$TenantId,
        [pscredential]$Credential,
        [string]$GroupId,
        [string]$ChunkPath
    )

    $ErrorActionPreference = 'Stop'
    Import-Module AzureAD -ErrorAction Stop

    Connect-AzureAD -TenantId $TenantId -Credential $Credential | Out-Null

    $sb = New-Object System.Text.StringBuilder

    try {
        $members = Get-AzureADGroupMember -ObjectId $GroupId -All $true

        foreach ($member in $members) {
            [void]$sb.Append($GroupId)
            [void]$sb.Append(',')
            [void]$sb.Append($member.ObjectId)
            [void]$sb.AppendLine()
        }

        [System.IO.File]::WriteAllText($ChunkPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))

        [pscustomobject]@{
            GroupId     = $GroupId
            Success     = $true
            MemberCount = @($members).Count
            ChunkPath   = $ChunkPath
            Error       = $null
        }
    }
    catch {
        # Write an empty chunk so merge logic stays simple
        [System.IO.File]::WriteAllText($ChunkPath, "", [System.Text.UTF8Encoding]::new($false))

        [pscustomobject]@{
            GroupId     = $GroupId
            Success     = $false
            MemberCount = 0
            ChunkPath   = $ChunkPath
            Error       = $_.Exception.Message
        }
    }
}

# ---------------------------
# Create runspace pool
# ---------------------------

$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$pool = [runspacefactory]::CreateRunspacePool(1, $ThreadCount, $iss, $Host)
$pool.Open()

$jobs = New-Object System.Collections.Generic.List[object]

Write-Host "Starting parallel member collection with $ThreadCount threads..." -ForegroundColor Cyan

$index = 0
foreach ($group in $groups) {
    $index++
    $chunkPath = Join-Path $tempDir ("chunk_{0:D8}.csv" -f $index)

    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool

    [void]$ps.AddScript($workerScript).
        AddArgument($TenantId).
        AddArgument($credential).
        AddArgument($group.Id).
        AddArgument($chunkPath)

    $handle = $ps.BeginInvoke()

    $jobs.Add([pscustomobject]@{
        PowerShell = $ps
        Handle     = $handle
        GroupId    = $group.Id
        ChunkPath  = $chunkPath
    }) | Out-Null
}

# ---------------------------
# Collect results as jobs finish
# ---------------------------

$completed = 0
$failed = New-Object System.Collections.Generic.List[object]

foreach ($job in $jobs) {
    try {
        $result = $job.PowerShell.EndInvoke($job.Handle)
        $completed++

        if (-not $result.Success) {
            $failed.Add($result) | Out-Null
        }

        if (($completed % 100) -eq 0 -or $completed -eq $groupCount) {
            Write-Host ("Completed {0}/{1} groups..." -f $completed, $groupCount) -ForegroundColor DarkCyan
        }
    }
    finally {
        $job.PowerShell.Dispose()
    }
}

$pool.Close()
$pool.Dispose()

# ---------------------------
# Merge chunk files into one CSV
# ---------------------------

Write-Host "Merging chunk files..." -ForegroundColor Cyan

Get-ChildItem -LiteralPath $tempDir -File | Sort-Object Name | ForEach-Object {
    if ($_.Length -gt 0) {
        Get-Content -LiteralPath $_.FullName -Encoding UTF8 | Add-Content -LiteralPath $tempCsv -Encoding UTF8
    }
}

# ---------------------------
# Convert CSV -> XLSX
# ---------------------------

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

# ---------------------------
# Failure summary
# ---------------------------

if ($failed.Count -gt 0) {
    $errorLog = Join-Path $tempRoot 'FailedGroups.csv'
    $failed | Select-Object GroupId, Error | Export-Csv -LiteralPath $errorLog -NoTypeInformation -Encoding UTF8
    Write-Warning "$($failed.Count) groups failed. Failure log: $errorLog"
}
else {
    Write-Host "All groups processed successfully." -ForegroundColor Green
}

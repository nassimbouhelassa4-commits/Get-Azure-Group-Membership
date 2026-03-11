param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsv,

    [Parameter(Mandatory = $true)]
    [string]$OutputCsv,

    [int]$ThrottleLimit = 8
)

# Requires PowerShell 7+
if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw "This script requires PowerShell 7 or later."
}

Write-Host "Loading input CSV..."
$rows = Import-Csv -Path $InputCsv -Delimiter ';'

if (-not $rows -or $rows.Count -eq 0) {
    throw "Input CSV is empty or could not be read."
}

# Validate required columns
$requiredColumns = @(
    'GroupId', 'GroupName', 'MemberId', 'MemberName',
    'MemberType', 'MemberUpn', 'TenantId'
)

$missingColumns = $requiredColumns | Where-Object {
    $_ -notin $rows[0].PSObject.Properties.Name
}

if ($missingColumns.Count -gt 0) {
    throw "Missing required columns: $($missingColumns -join ', ')"
}

Write-Host "Building in-memory indexes..."

# Direct membership by parent group
# Hashtable: GroupId -> direct membership rows
$membersByGroup = @{}
foreach ($row in $rows) {
    if (-not $membersByGroup.ContainsKey($row.GroupId)) {
        $membersByGroup[$row.GroupId] = [System.Collections.Generic.List[object]]::new()
    }
    $membersByGroup[$row.GroupId].Add($row)
}

# List of unique root groups from the source file
$rootGroups = $rows |
    Select-Object GroupId, GroupName, TenantId -Unique

Write-Host "Expanding nested memberships with multithreading..."

$expanded = $rootGroups | ForEach-Object -Parallel {
    $rootGroup = $_
    $membersIndex = $using:membersByGroup

    $localResults = [System.Collections.Generic.List[object]]::new()

    # Stack for DFS
    $stack = [System.Collections.Generic.Stack[object]]::new()

    # Track visited nested groups for this root group to avoid cycles
    $visitedGroups = [System.Collections.Generic.HashSet[string]]::new()

    # Track emitted rows to avoid duplicates
    # Key format: RootGroupId|MemberId
    $emitted = [System.Collections.Generic.HashSet[string]]::new()

    # Seed with direct members of the root group
    if ($membersIndex.ContainsKey($rootGroup.GroupId)) {
        foreach ($directMember in $membersIndex[$rootGroup.GroupId]) {
            $stack.Push([PSCustomObject]@{
                RootGroupId   = $rootGroup.GroupId
                RootGroupName = $rootGroup.GroupName
                RootTenantId  = $rootGroup.TenantId
                CurrentRow    = $directMember
            })
        }
    }

    while ($stack.Count -gt 0) {
        $item = $stack.Pop()
        $row  = $item.CurrentRow

        $memberKey = "$($item.RootGroupId)|$($row.MemberId)"
        if (-not $emitted.Contains($memberKey)) {
            $null = $emitted.Add($memberKey)

            # Output row with the SAME original columns
            $localResults.Add([PSCustomObject]@{
                GroupId     = $item.RootGroupId
                GroupName   = $item.RootGroupName
                MemberId    = $row.MemberId
                MemberName  = $row.MemberName
                MemberType  = $row.MemberType
                MemberUpn   = $row.MemberUpn
                TenantId    = $item.RootTenantId
            })
        }

        # If the member is itself a group, recurse into it
        $isNestedGroup = $false
        if ($row.MemberType) {
            $normalizedType = $row.MemberType.ToString().Trim().ToLowerInvariant()
            if ($normalizedType -match 'group') {
                $isNestedGroup = $true
            }
        }

        if ($isNestedGroup -and $row.MemberId) {
            if (-not $visitedGroups.Contains($row.MemberId)) {
                $null = $visitedGroups.Add($row.MemberId)

                if ($membersIndex.ContainsKey($row.MemberId)) {
                    foreach ($childMember in $membersIndex[$row.MemberId]) {
                        $stack.Push([PSCustomObject]@{
                            RootGroupId   = $item.RootGroupId
                            RootGroupName = $item.RootGroupName
                            RootTenantId  = $item.RootTenantId
                            CurrentRow    = $childMember
                        })
                    }
                }
            }
        }
    }

    $localResults
} -ThrottleLimit $ThrottleLimit

Write-Host "Writing output CSV..."
$expanded |
    Sort-Object GroupName, GroupId, MemberType, MemberName, MemberId |
    Export-Csv -Path $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8

Write-Host "Done. Output written to: $OutputCsv"

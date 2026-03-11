param(
    [Parameter(Mandatory = $false)]
    [string]$OutputCsvPath = ".\EntraID_GroupMemberships.csv"
)

# Requires AzureAD module
# Install-Module AzureAD
# Import-Module AzureAD

try {
    if (-not (Get-Module -ListAvailable -Name AzureAD)) {
        throw "AzureAD module is not installed. Install it first with: Install-Module AzureAD"
    }

    Import-Module AzureAD -ErrorAction Stop

    # Connect if needed
    try {
        $null = Get-AzureADTenantDetail -ErrorAction Stop
    }
    catch {
        Connect-AzureAD -ErrorAction Stop
    }

    $tenantDetail = Get-AzureADTenantDetail -ErrorAction Stop
    $tenantId = $tenantDetail.ObjectId

    $results = New-Object System.Collections.Generic.List[object]

    function Get-MemberDetails {
        param(
            [Parameter(Mandatory = $true)]
            $Member
        )

        $memberId   = $Member.ObjectId
        $memberType = $Member.ObjectType
        $memberName = $null
        $memberUpn  = $null

        switch ($memberType) {
            "User" {
                try {
                    $user = Get-AzureADUser -ObjectId $memberId -ErrorAction Stop
                    $memberName = $user.DisplayName
                    $memberUpn  = $user.UserPrincipalName
                }
                catch {
                    $memberName = $Member.DisplayName
                    $memberUpn  = $Member.UserPrincipalName
                }
            }

            "Group" {
                try {
                    $group = Get-AzureADGroup -ObjectId $memberId -ErrorAction Stop
                    $memberName = $group.DisplayName
                    $memberUpn  = $null
                }
                catch {
                    $memberName = $Member.DisplayName
                    $memberUpn  = $null
                }
            }

            "ServicePrincipal" {
                try {
                    $sp = Get-AzureADServicePrincipal -ObjectId $memberId -ErrorAction Stop
                    $memberName = $sp.DisplayName
                    $memberUpn  = $null
                }
                catch {
                    $memberName = $Member.DisplayName
                    $memberUpn  = $null
                }
            }

            "Device" {
                try {
                    $device = Get-AzureADDevice -ObjectId $memberId -ErrorAction Stop
                    $memberName = $device.DisplayName
                    $memberUpn  = $null
                }
                catch {
                    $memberName = $Member.DisplayName
                    $memberUpn  = $null
                }
            }

            "Contact" {
                try {
                    $contact = Get-AzureADContact -ObjectId $memberId -ErrorAction Stop
                    $memberName = $contact.DisplayName
                    $memberUpn  = $contact.Mail
                }
                catch {
                    $memberName = $Member.DisplayName
                    $memberUpn  = $null
                }
            }

            default {
                $memberName = $Member.DisplayName
                $memberUpn  = $null
            }
        }

        [PSCustomObject]@{
            MemberId   = $memberId
            MemberName = $memberName
            MemberUpn  = $memberUpn
            MemberType = $memberType
        }
    }

    $groups = Get-AzureADGroup -All $true

    $groupCount = $groups.Count
    $currentGroup = 0

    foreach ($group in $groups) {
        $currentGroup++

        Write-Progress -Activity "Reading Azure Entra ID groups" `
                       -Status "Processing group $currentGroup of $groupCount : $($group.DisplayName)" `
                       -PercentComplete (($currentGroup / $groupCount) * 100)

        try {
            $members = Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true -ErrorAction Stop

            if (-not $members -or $members.Count -eq 0) {
                continue
            }

            foreach ($member in $members) {
                $memberDetails = Get-MemberDetails -Member $member

                $results.Add([PSCustomObject]@{
                    TenantId   = $tenantId
                    GroupId    = $group.ObjectId
                    FroupName  = $group.DisplayName   # kept exactly as requested
                    MemberId   = $memberDetails.MemberId
                    MemberName = $memberDetails.MemberName
                    MemberUpn  = $memberDetails.MemberUpn
                    MemberType = $memberDetails.MemberType
                })
            }
        }
        catch {
            Write-Warning "Failed to read members for group '$($group.DisplayName)' ($($group.ObjectId)): $($_.Exception.Message)"
        }
    }

    $results | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

    Write-Host "Export completed: $OutputCsvPath" -ForegroundColor Green
    Write-Host "Rows exported: $($results.Count)" -ForegroundColor Green
}
catch {
    Write-Error $_.Exception.Message
}

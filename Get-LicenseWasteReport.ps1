<#
.SYNOPSIS
    Finds Microsoft 365 license spend you can reclaim - read-only.

.DESCRIPTION
    Sweeps the tenant for the three classic license leaks:

      1. UNASSIGNED SEATS  - purchased seats no one is using, per SKU
      2. DISABLED USERS    - sign-in-disabled accounts still holding licenses
      3. STALE USERS       - enabled accounts holding licenses that haven't
                             signed in for -StaleDays days (or ever)

    Prints a per-SKU summary and a reclaim-candidate list, returns structured
    objects for further processing, and optionally exports the per-user detail
    to CSV for the renewal conversation.

    Makes no changes to the tenant. Safe to run any time.

.PARAMETER StaleDays
    Days without an interactive sign-in before a licensed, enabled user is
    counted as stale. Default: 90.

.PARAMETER CsvPath
    Optional path for a CSV export of the per-user reclaim candidates.

.EXAMPLE
    .\Get-LicenseWasteReport.ps1

    Prints the waste summary for the tenant using the 90-day stale threshold.

.EXAMPLE
    .\Get-LicenseWasteReport.ps1 -StaleDays 60 -CsvPath .\license-waste.csv

    Uses a 60-day threshold and writes the reclaim-candidate detail to CSV.

.NOTES
    Required Graph scopes (all read-only):
        User.Read.All, Organization.Read.All, AuditLog.Read.All
    Required modules:
        Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement

    SignInActivity requires an Entra ID P1/P2 license in the tenant; without
    it the stale-user check is skipped with a warning.

    Sign-in data can lag by up to 24 hours, and 'stale' is a conversation
    starter, not a verdict - check with the user's manager before reclaiming.
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 3650)][int]$StaleDays = 90,
    [string]$CsvPath
)

$ErrorActionPreference = 'Stop'

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes 'User.Read.All','Organization.Read.All','AuditLog.Read.All' -NoWelcome
}

#region Collect
Write-Host 'Collecting subscribed SKUs...'
$skus = Get-MgSubscribedSku
$skuNameById = @{}
foreach ($s in $skus) { $skuNameById[$s.SkuId] = $s.SkuPartNumber }

Write-Host 'Collecting licensed users (this can take a while in large tenants)...'
$licensedUsers = Get-MgUser -All -ConsistencyLevel eventual -CountVariable licensedCount `
    -Filter "assignedLicenses/`$count ne 0" `
    -Property Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses,SignInActivity

$signInDataAvailable = $true
$probe = $licensedUsers | Select-Object -First 1
if ($probe -and -not $probe.SignInActivity) {
    # SignInActivity comes back null on every user when the tenant lacks Entra P1/P2.
    $nullCount = @($licensedUsers | Where-Object { -not $_.SignInActivity }).Count
    if ($nullCount -eq @($licensedUsers).Count) {
        $signInDataAvailable = $false
        Write-Warning 'SignInActivity is empty for all users (Entra P1/P2 may be missing) - stale-user check skipped.'
    }
}
#endregion

#region 1. Unassigned seats per SKU
$skuRows = foreach ($s in $skus) {
    $purchased = $s.PrepaidUnits.Enabled
    $suspended = $s.PrepaidUnits.Suspended + $s.PrepaidUnits.Warning
    [pscustomobject]@{
        Sku        = $s.SkuPartNumber
        Purchased  = $purchased
        Assigned   = $s.ConsumedUnits
        Unassigned = [Math]::Max(0, $purchased - $s.ConsumedUnits)
        SuspendedOrWarning = $suspended
    }
}
#endregion

#region 2 + 3. Per-user reclaim candidates
$cutoff = (Get-Date).AddDays(-$StaleDays)
$reclaimCandidates = foreach ($u in $licensedUsers) {
    $userSkus = (@($u.AssignedLicenses) | ForEach-Object {
        $id = $_.SkuId
        if ($skuNameById.ContainsKey($id)) { $skuNameById[$id] } else { $id }
    }) -join '; '

    if (-not $u.AccountEnabled) {
        [pscustomobject]@{
            Reason            = 'DISABLED ACCOUNT'
            DisplayName       = $u.DisplayName
            UserPrincipalName = $u.UserPrincipalName
            LastSignIn        = $u.SignInActivity.LastSignInDateTime
            Licenses          = $userSkus
        }
        continue
    }

    if ($signInDataAvailable) {
        $lastSignIn = $u.SignInActivity.LastSignInDateTime
        if (-not $lastSignIn -or $lastSignIn -lt $cutoff) {
            [pscustomobject]@{
                Reason            = if ($lastSignIn) { "STALE (> $StaleDays days)" } else { 'NEVER SIGNED IN' }
                DisplayName       = $u.DisplayName
                UserPrincipalName = $u.UserPrincipalName
                LastSignIn        = $lastSignIn
                Licenses          = $userSkus
            }
        }
    }
}
$reclaimCandidates = @($reclaimCandidates)
#endregion

#region Output
Write-Host ''
Write-Host '===================== LICENSE WASTE REPORT ====================='
Write-Host ("  Licensed users: {0}    Stale threshold: {1} days" -f @($licensedUsers).Count, $StaleDays)
Write-Host ''
Write-Host '  --- Seats purchased vs assigned (per SKU) ---'
$skuRows | Sort-Object Unassigned -Descending | ForEach-Object {
    $flag = if ($_.Unassigned -gt 0) { ' <-- unassigned seats' } else { '' }
    Write-Host ("    {0}: {1} purchased / {2} assigned / {3} unassigned{4}" -f $_.Sku, $_.Purchased, $_.Assigned, $_.Unassigned, $flag)
}
Write-Host ''
$disabled = @($reclaimCandidates | Where-Object Reason -eq 'DISABLED ACCOUNT')
$stale    = @($reclaimCandidates | Where-Object Reason -ne 'DISABLED ACCOUNT')
Write-Host ("  --- Licensed but disabled accounts: {0} ---" -f $disabled.Count)
$disabled | ForEach-Object { Write-Host ("    {0} <{1}>  [{2}]" -f $_.DisplayName, $_.UserPrincipalName, $_.Licenses) }
if ($signInDataAvailable) {
    Write-Host ''
    Write-Host ("  --- Licensed, enabled, but stale/never signed in: {0} ---" -f $stale.Count)
    $stale | ForEach-Object { Write-Host ("    {0} <{1}>  last: {2}  [{3}]" -f $_.DisplayName, $_.UserPrincipalName, $(if ($_.LastSignIn) { $_.LastSignIn } else { 'never' }), $_.Licenses) }
}
Write-Host ''
$totalUnassigned = ($skuRows | Measure-Object Unassigned -Sum).Sum
Write-Host ("  TOTALS: {0} unassigned seat(s), {1} license-holding disabled account(s), {2} stale candidate(s)" -f $totalUnassigned, $disabled.Count, $stale.Count)
Write-Host '  Multiply by your per-seat prices for the renewal conversation.'
Write-Host '================================================================='
Write-Host ''

if ($CsvPath) {
    $reclaimCandidates | Export-Csv -Path $CsvPath -NoTypeInformation
    Write-Host "Per-user detail exported to: $CsvPath"
}

# Structured result for pipelines
[pscustomobject]@{
    GeneratedUtc      = (Get-Date).ToUniversalTime().ToString('o')
    StaleDays         = $StaleDays
    SkuSummary        = $skuRows
    ReclaimCandidates = $reclaimCandidates
}
#endregion

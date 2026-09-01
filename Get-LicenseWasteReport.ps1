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

.PARAMETER JsonPath
    Optional path for a JSON export of the whole result - per-SKU summary,
    consumption SKUs, and reclaim candidates. This is the machine-readable
    feed (it-ops-console reads it).

.PARAMETER ConsumptionSkuThreshold
    SKUs with this many (or more) prepaid units are treated as self-service /
    consumption / viral SKUs (Power Apps Dev, Flow Free, PSTN consumption,
    Forms Pro trials...) - Microsoft reports those with 10,000+ fake "seats",
    which would drown the real waste numbers. They are listed separately and
    excluded from the totals. Default: 10000. Raise it if your org genuinely
    buys 10k+ seats of something.

.PARAMETER PriceList
    Optional path to an INI of per-seat monthly prices, so the report can put
    a dollar figure on the waste instead of just a seat count. Format:

        [settings]
        currency = $           ; prefix shown before amounts (default: $)
        [prices]
        SPB            = 33.00 ; the SKU name on the left is the SkuPartNumber
        AAD_PREMIUM_P2 = 9.00  ; this report already prints in its SKU summary

    Only the SKUs you price get a dollar figure; the rest still show their
    seat counts. From the priced SKUs the report computes two numbers:
    money sitting in unassigned seats, and money tied up in the licenses that
    disabled or stale accounts still hold (reclaimable now). With no
    -PriceList, nothing about money appears and everything else is unchanged.

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
    [string]$CsvPath,
    [string]$JsonPath,
    [ValidateRange(100, 100000000)][int]$ConsumptionSkuThreshold = 10000,
    [string]$PriceList
)

$ErrorActionPreference = 'Stop'

function Read-PriceList {
    # Tiny INI reader for the price file: [prices] SKU = number, plus an
    # optional [settings] currency = <prefix>. Blank/comment (# or ;) lines and
    # non-numeric values are skipped, so a half-filled template is harmless.
    param([string]$Path)
    $out = [ordered]@{ Currency = '$'; Prices = @{} }
    if (-not $Path -or -not (Test-Path $Path)) { return $out }
    $section = ''
    foreach ($raw in Get-Content -LiteralPath $Path) {
        $line = $raw.Trim()
        if ($line -eq '' -or $line[0] -eq '#' -or $line[0] -eq ';') { continue }
        if ($line -match '^\[(.+)\]$') { $section = $Matches[1].Trim().ToLowerInvariant(); continue }
        $eq = $line.IndexOf('=')
        if ($eq -lt 1) { continue }
        $key = $line.Substring(0, $eq).Trim()
        $val = $line.Substring($eq + 1).Trim()
        # Drop an inline comment after the value (e.g. "33.00 ; E3").
        $val = ($val -split '\s+[;#]', 2)[0].Trim()
        if ($section -eq 'settings' -and $key.ToLowerInvariant() -eq 'currency') {
            if ($val) { $out.Currency = $val }
        }
        elseif ($section -eq 'prices') {
            $num = 0.0
            if ([double]::TryParse($val, [ref]$num)) { $out.Prices[$key] = $num }
        }
    }
    return $out
}

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
        IsConsumption      = ($purchased -ge $ConsumptionSkuThreshold)
    }
}
$countedSkuRows     = @($skuRows | Where-Object { -not $_.IsConsumption })
$consumptionSkuRows = @($skuRows | Where-Object { $_.IsConsumption })
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
$countedSkuRows | Sort-Object Unassigned -Descending | ForEach-Object {
    $flag = if ($_.Unassigned -gt 0) { ' <-- unassigned seats' } else { '' }
    Write-Host ("    {0}: {1} purchased / {2} assigned / {3} unassigned{4}" -f $_.Sku, $_.Purchased, $_.Assigned, $_.Unassigned, $flag)
}
if ($consumptionSkuRows.Count) {
    Write-Host ''
    Write-Host ("  --- Self-service / consumption SKUs (excluded from totals; threshold {0}) ---" -f $ConsumptionSkuThreshold)
    $consumptionSkuRows | Sort-Object Sku | ForEach-Object {
        Write-Host ("    {0}: {1} in use" -f $_.Sku, $_.Assigned)
    }
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
$totalUnassigned = ($countedSkuRows | Measure-Object Unassigned -Sum).Sum
if ($null -eq $totalUnassigned) { $totalUnassigned = 0 }
Write-Host ("  TOTALS: {0} unassigned seat(s), {1} license-holding disabled account(s), {2} stale candidate(s)" -f $totalUnassigned, $disabled.Count, $stale.Count)

# --- Costing: put a dollar figure on the waste, if a price list was given --- #
# Only SKUs with a price get a figure; the report multiplies here so the
# renewal conversation does not have to. Two numbers: money in unassigned
# seats, and money in the licenses that disabled/stale accounts still hold.
$priceListProvided = [bool]$PriceList -and (Test-Path $PriceList)
$priceData = Read-PriceList $PriceList
$prices = $priceData.Prices
$costing = $null
if ($priceListProvided) {
    $pricedCount = 0
    foreach ($row in $countedSkuRows) {
        $price = if ($prices.ContainsKey($row.Sku)) { [double]$prices[$row.Sku] } else { $null }
        if ($null -ne $price) { $pricedCount++ }
        $unusedCost = if ($null -ne $price) { [math]::Round($price * [int]$row.Unassigned, 2) } else { $null }
        $row | Add-Member -NotePropertyName MonthlyPrice -NotePropertyValue $price -Force
        $row | Add-Member -NotePropertyName UnusedMonthlyCost -NotePropertyValue $unusedCost -Force
    }
    foreach ($c in $reclaimCandidates) {
        $skuList = @("$($c.Licenses)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $sum = 0.0; $anyPriced = $false
        foreach ($s in $skuList) { if ($prices.ContainsKey($s)) { $sum += [double]$prices[$s]; $anyPriced = $true } }
        $cost = if ($anyPriced) { [math]::Round($sum, 2) } else { $null }
        $c | Add-Member -NotePropertyName MonthlyCost -NotePropertyValue $cost -Force
    }
    $unusedMonthly = ($countedSkuRows | Where-Object { $null -ne $_.UnusedMonthlyCost } | Measure-Object UnusedMonthlyCost -Sum).Sum
    if ($null -eq $unusedMonthly) { $unusedMonthly = 0 }
    $reclaimMonthly = ($reclaimCandidates | Where-Object { $null -ne $_.MonthlyCost } | Measure-Object MonthlyCost -Sum).Sum
    if ($null -eq $reclaimMonthly) { $reclaimMonthly = 0 }
    $unpricedSkus = @($countedSkuRows | Where-Object { $null -eq $_.MonthlyPrice } | ForEach-Object { $_.Sku })
    $costing = [ordered]@{
        Currency           = $priceData.Currency
        HasPrices          = ($pricedCount -gt 0)
        PricedSkuCount     = $pricedCount
        UnpricedSkuCount   = @($unpricedSkus).Count
        UnpricedSkus       = $unpricedSkus
        UnusedSeatsMonthly = [math]::Round($unusedMonthly, 2)
        UnusedSeatsAnnual  = [math]::Round($unusedMonthly * 12, 2)
        ReclaimableMonthly = [math]::Round($reclaimMonthly, 2)
        ReclaimableAnnual  = [math]::Round($reclaimMonthly * 12, 2)
    }
    $cur = $priceData.Currency
    if ($costing.HasPrices) {
        Write-Host ("  IN MONEY: {0}{1:N0}/mo in unassigned seats ({0}{2:N0}/yr); {0}{3:N0}/mo reclaimable from disabled/stale accounts ({0}{4:N0}/yr)." -f `
            $cur, $costing.UnusedSeatsMonthly, $costing.UnusedSeatsAnnual, $costing.ReclaimableMonthly, $costing.ReclaimableAnnual)
        if ($costing.UnpricedSkuCount -gt 0) {
            Write-Host ("  ({0} SKU(s) have no price yet: {1})" -f $costing.UnpricedSkuCount, ($unpricedSkus -join ', '))
        }
    } else {
        Write-Host "  Price list found but no prices set yet - add per-seat prices to $PriceList and re-run for dollar figures."
    }
} else {
    Write-Host '  Multiply by your per-seat prices for the renewal conversation (or pass -PriceList for the dollars).'
}
Write-Host '================================================================='
Write-Host ''

if ($CsvPath) {
    $reclaimCandidates | Export-Csv -Path $CsvPath -NoTypeInformation
    Write-Host "Per-user detail exported to: $CsvPath"
}

# Structured result for pipelines
$result = [pscustomobject]@{
    GeneratedUtc      = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    StaleDays         = $StaleDays
    LicensedUsers     = @($licensedUsers).Count
    SkuSummary        = $countedSkuRows
    ConsumptionSkus   = $consumptionSkuRows
    ReclaimCandidates = $reclaimCandidates
    Costing           = $costing
}

if ($JsonPath) {
    $result | ConvertTo-Json -Depth 6 | Set-Content -Path $JsonPath -Encoding UTF8
    Write-Host "Full result exported to: $JsonPath"
}

$result
#endregion

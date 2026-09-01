# M365 License Waste Report

One read-only PowerShell script that finds the Microsoft 365 license spend you can reclaim, in the three places it always hides:

1. **Unassigned seats** — SKUs you're paying for that nobody is using
2. **Disabled accounts** — sign-in-disabled users still holding licenses
3. **Stale users** — enabled, licensed accounts that haven't signed in for 90+ days (threshold adjustable)

```powershell
.\Get-LicenseWasteReport.ps1 -StaleDays 60 -CsvPath .\license-waste.csv
```

```
===================== LICENSE WASTE REPORT =====================
  Licensed users: 187    Stale threshold: 60 days

  --- Seats purchased vs assigned (per SKU) ---
    O365_BUSINESS_PREMIUM: 150 purchased / 141 assigned / 9 unassigned <-- unassigned seats
    EMS: 150 purchased / 150 assigned / 0 unassigned

  --- Licensed but disabled accounts: 3 ---
    ...

  --- Licensed, enabled, but stale/never signed in: 5 ---
    ...

  TOTALS: 9 unassigned seat(s), 3 license-holding disabled account(s), 5 stale candidate(s)
  Multiply by your per-seat prices for the renewal conversation.
=================================================================
```

The script prints the summary, returns a structured object (pipe it wherever), and `-CsvPath` exports the per-user reclaim candidates for the renewal conversation with your reseller.

## Put a dollar figure on it

Hand it a price list and it does the multiplication for you:

```powershell
.\Get-LicenseWasteReport.ps1 -PriceList .\prices.ini -JsonPath .\licensing.json
```

`prices.ini` is per-seat **monthly** prices keyed by the SKU part number the report already prints (copy `prices.example.ini` to start):

```ini
[settings]
currency = $
[prices]
SPB            = 33.00
AAD_PREMIUM_P2 = 9.00
```

You get two numbers: money sitting in **unassigned seats**, and money tied up in the licenses that **disabled or stale accounts** still hold (reclaimable now) — each as a monthly and an annual figure. Only the SKUs you price get a dollar figure; the rest keep showing seat counts. Skip `-PriceList` and nothing about money appears. (In the full suite you don't usually edit this file by hand — after your first refresh, the console lists your tenant's real SKUs in `prices.ini` for you to fill in.)

## Setup

```powershell
Install-Module Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
Connect-MgGraph -Scopes "User.Read.All","Organization.Read.All","AuditLog.Read.All"   # all read-only
```

The script connects on its own if you haven't.

## Notes and caveats

- **Read-only.** The script changes nothing — it reports. Reclaiming is your call.
- **Run against a real production tenant** (200+ licensed users) — the numbers below are the shape of what it finds.
- **Self-service and consumption SKUs are separated out.** Microsoft reports Power Apps Dev, Flow Free, PSTN consumption and similar with 10,000 — or 10,000,000 — nominal "seats". Left in the math they produce nonsense totals like "13 million unassigned seats", so anything at or above `-ConsumptionSkuThreshold` (default 10,000) is listed in its own section and excluded from the totals. Raise the threshold if your org genuinely buys 10k+ seats of something.
- Last-sign-in data (`SignInActivity`) requires Entra ID P1/P2 in the tenant. Without it, the stale-user check is skipped with a warning; the other two checks still work.
- Sign-in timestamps can lag by up to 24 hours, and "stale" is a conversation starter, not a verdict — a licensed account with no sign-ins might be a service account doing its job. Check before you cut.
- Disabled accounts holding licenses aren't always waste either — a mailbox mid-conversion to shared still needs its license briefly. See [entra-lifecycle-toolkit](https://github.com/JonathanT10/entra-lifecycle-toolkit) for the offboarding flow that avoids that trap.

## License

MIT — see [LICENSE](LICENSE).

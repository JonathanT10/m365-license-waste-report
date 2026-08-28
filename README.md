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

## Setup

```powershell
Install-Module Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
Connect-MgGraph -Scopes "User.Read.All","Organization.Read.All","AuditLog.Read.All"   # all read-only
```

The script connects on its own if you haven't.

## Notes and caveats

- **Read-only.** The script changes nothing — it reports. Reclaiming is your call.
- Last-sign-in data (`SignInActivity`) requires Entra ID P1/P2 in the tenant. Without it, the stale-user check is skipped with a warning; the other two checks still work.
- Sign-in timestamps can lag by up to 24 hours, and "stale" is a conversation starter, not a verdict — a licensed account with no sign-ins might be a service account doing its job. Check before you cut.
- Disabled accounts holding licenses aren't always waste either — a mailbox mid-conversion to shared still needs its license briefly. See [entra-lifecycle-toolkit](https://github.com/JonathanT10/entra-lifecycle-toolkit) for the offboarding flow that avoids that trap.

## License

MIT — see [LICENSE](LICENSE).

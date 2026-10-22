# Copilot Instructions

This repo contains a single Lantern employee-offboarding runbook in `offboarding/`, implemented twice:

- `offboarding-exchange-actions.ps1` – original Azure Automation runbook (PowerShell 7.2, Microsoft Graph + ExchangeOnlineManagement modules).
- `invoke_offboarding.py` – Python port of the same runbook (`requests` + `azure-identity`).

There is no build system, test suite, or linter configured.

## Running

```powershell
# Python (deps: pip install -r offboarding\requirements.txt; also needs `pwsh` with the ExchangeOnlineManagement module)
python offboarding\invoke_offboarding.py --user-principal-name user@lanternstudios.com --last-day 2026-10-22 --what-if

# PowerShell (Azure Automation; -WhatIf for dry run)
.\offboarding\offboarding-exchange-actions.ps1 -UserPrincipalName user@lanternstudios.com -LastDay 2026-10-22 -WhatIf
```

Both authenticate with a **managed identity**, so they only work end-to-end inside Azure. Always validate with `--what-if` / `-WhatIf`; never run against a real user unless asked. The Freshservice key comes from `FRESHSERVICE_API_KEY` (Python) or the encrypted Automation variable `FreshserviceApiKey` (PowerShell).

## Architecture

Both scripts run the same ordered steps, and **the order is load-bearing**: licensing is driven by dynamic groups keyed on Entra attributes, so attributes may only be cleared after the mailbox is confirmed Shared.

1. Guard: refuse users holding directory roles
2. Disable account
3. Exchange: convert to Shared, hide from GAL, forwarding + auto-reply
4. Gate: poll until mailbox is `SharedMailbox` (30 × 20s), abort otherwise; warn on size ≥ 50 GB or litigation hold
5. Clear Entra attributes, remove manager, remove non-dynamic group memberships
6. Intune device category (`Lantern-In Transit` if laptop leaves with user; `Lantern-retired` if left in office and device ≥3 years old; else `Lantern-Inventory`)
7. Laptop-return email as a reply on the Freshservice ticket (not mailed from helpdesk@, which would open a new ticket)

Exchange Online-only settings have no Graph REST API, so the Python port shells out to `pwsh` (`run_pwsh`) for those cmdlets and uses Graph directly for everything else.

## Conventions

- **Keep the two implementations in sync**: any change to behavior, step order, strings (emails, category names), or defaults (`RetentionWeeks=8`, tenant/Freshservice domains) must be applied to both files.
- Every mutating action must support dry-run: PowerShell uses `SupportsShouldProcess`/`$PSCmdlet.ShouldProcess`; Python wraps mutations in `if not what_if` and logs `[WhatIf] Would ...`. Step 4's verification gate is skipped in what-if mode.
- Errors are fatal (`$ErrorActionPreference = 'Stop'` / top-level exception → exit 1), except per-group removal failures, which only warn.
- Dynamic-membership groups are skipped when removing groups; they drop automatically after the attribute change.
- Intune category assignment uses Graph `/beta`; most else uses `/v1.0`.
- The script only logs a reminder date (last day + retention weeks) to delete the shared mailbox; it does not delete it.
- Required Graph permissions and Exchange role are documented in each script's header; update the header when adding calls that need new permissions.

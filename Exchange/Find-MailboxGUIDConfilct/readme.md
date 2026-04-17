# Find-MailboxGuidConflict

A PowerShell diagnostic script for resolving **"potential mailbox guid conflict"** errors during cross-tenant mailbox migrations in Microsoft 365.

---

## The Problem

When migrating mailboxes between tenants, the Microsoft Exchange Migration Service (MRS) can fail with:

```
MRSRemotePermanentException: Recipient's GUID '00000000-...' does not match the expected value
'c4beabc8-b4e6-44da-8196-e6c1dd431093'. This issue could occur if mailbox guid
'c4beabc8-b4e6-44da-8196-e6c1dd431093' is stamped on multiple recipients.
```

This happens when an `ExchangeGuid` is found on more than one recipient object — including **soft-deleted objects** and **inactive mailboxes** that are invisible to standard searches in the Exchange Admin Center or basic PowerShell queries.

### Hold-retained inactive mailboxes

A particularly sneaky cause occurs when the **source tenant uses Litigation Hold or In-Place Hold**. When either hold type is active, Exchange retains the mailbox as an *inactive mailbox* even after the user account is removed. This inactive mailbox keeps the original `ExchangeGuid`, making the conflict invisible until you check for it explicitly — which is exactly what this script does.

> **Before investigating:** Check whether the affected mailboxes have holds applied:
> ```powershell
> Get-Mailbox <alias> | Select-Object DisplayName, LitigationHoldEnabled, InPlaceHolds
> ```

---

## What This Script Does

Sweeps **all recipient types** in the connected tenant for a matching `ExchangeGuid`:

| Recipient Type | Notes |
|---|---|
| Active Mailboxes | Standard user, shared, room, equipment |
| Soft-Deleted Mailboxes | Recoverable deleted mailboxes |
| **Inactive Mailboxes** | ⚠️ Hold-retained mailboxes — invisible to normal searches |
| Active Mail Users | MailUser objects (common in target tenant) |
| **Soft-Deleted Mail Users** | ⚠️ Most commonly missed — frequent root cause |
| Mail Contacts | External contacts with ExchangeGuid stamps |
| Remote Mailboxes | Hybrid Exchange environments only |
| Distribution Groups | Mail-enabled groups |
| Microsoft 365 Groups | Unified Groups |

---

## Requirements

- PowerShell 5.1 or PowerShell 7+
- [ExchangeOnlineManagement module](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
- **Exchange Administrator** or **Global Administrator** role in the tenant being scanned

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

## Usage

### Basic scan

```powershell
# 1. Connect to your tenant
Connect-ExchangeOnline -UserPrincipalName admin@yourtenant.onmicrosoft.com

# 2. Run the script with the GUID from the error message
.\Find-MailboxGuidConflict.ps1 -ConflictGuid 'c4beabc8-b4e6-44da-8196-e6c1dd431093'
```

### Scan with CSV export

```powershell
.\Find-MailboxGuidConflict.ps1 -ConflictGuid 'c4beabc8-b4e6-44da-8196-e6c1dd431093' -ExportCsv 'C:\Temp\GuidConflictResults.csv'
```

### Full cross-tenant workflow

```powershell
# Step 1 — Scan SOURCE tenant
Connect-ExchangeOnline -UserPrincipalName admin@sourcetenant.onmicrosoft.com
.\Find-MailboxGuidConflict.ps1 -ConflictGuid 'c4beabc8-b4e6-44da-8196-e6c1dd431093'
Disconnect-ExchangeOnline

# Step 2 — Scan TARGET tenant
Connect-ExchangeOnline -UserPrincipalName admin@targettenant.onmicrosoft.com
.\Find-MailboxGuidConflict.ps1 -ConflictGuid 'c4beabc8-b4e6-44da-8196-e6c1dd431093'
Disconnect-ExchangeOnline
```

---

## Example Output

```
────────────────────────────────────────────────────────────
  Find-MailboxGuidConflict.ps1
  Cross-Tenant Migration Diagnostic Tool
────────────────────────────────────────────────────────────
  Target GUID : c4beabc8-b4e6-44da-8196-e6c1dd431093
  Tenant      : Contoso Ltd
────────────────────────────────────────────────────────────

  Checking Active Mailboxes... No match
  Checking Soft-Deleted Mailboxes... No match
  Checking Active Mail Users... No match
  Checking Soft-Deleted Mail Users... MATCH FOUND        ← common culprit
  Checking Mail Contacts... No match
  Checking Remote Mailboxes (Hybrid)... Skipped
  Checking Distribution Groups... No match
  Checking Microsoft 365 Groups... No match

────────────────────────────────────────────────────────────
  RESULTS SUMMARY
────────────────────────────────────────────────────────────

  1 object(s) found:

  ObjectType              DisplayName   PrimarySmtpAddress      RecipientTypeDetail
  ----------              -----------   ------------------      -------------------
  Soft-Deleted Mail Users John Doe      john.doe@contoso.com    MailUser
```

---

## Resolving the Conflict

Once you've identified the duplicate object, choose the appropriate fix:

**Duplicate MailUser — clear the GUID:**
```powershell
Set-MailUser john.doe -ExchangeGuid '00000000-0000-0000-0000-000000000000'
```

**Soft-deleted object — permanently remove it:**
```powershell
# Find the soft-deleted user
Get-MailUser -SoftDeletedMailUser | Where-Object { $_.Alias -eq 'john.doe' }

# Permanently delete via the compliance portal or:
Remove-MailUser john.doe -PermanentlyDelete
```

**Ghost object from a failed migration — clean up MigrationUser:**
```powershell
Remove-MigrationUser john.doe@sourcetenant.com
```

**Inactive mailbox retained by Litigation Hold or In-Place Hold:**
```powershell
# Step 1 — Check what holds are applied
Get-Mailbox john.doe | Select-Object LitigationHoldEnabled, InPlaceHolds

# Step 2 — Release Litigation Hold
Set-Mailbox john.doe -LitigationHoldEnabled $false

# Step 3 — Remove In-Place Holds via the Exchange Admin Center (classic EAC)
#           Compliance > In-Place eDiscovery & Hold > remove the hold entry

# Step 4 — After holds clear, the inactive mailbox becomes visible
Get-Mailbox -InactiveMailboxOnly | Where-Object { $_.Alias -eq 'john.doe' }

# Step 5 — Permanently delete the inactive mailbox
Remove-Mailbox john.doe -InactiveMailbox -PermanentlyDelete
```

> ⚠️ **Important:** Releasing a hold is a compliance action. Confirm with your legal or compliance team before removing any holds, as this may affect eDiscovery obligations.

After resolving, retry the migration batch.

---

## Common Root Causes

| Scenario | Tenant | Fix |
|---|---|---|
| MailUser pre-stamped with GUID | Target | Clear ExchangeGuid with `Set-MailUser` |
| Soft-deleted MailUser not cleaned up | Target | Permanently delete the object |
| Previous failed migration left a ghost object | Either | Remove-MigrationUser + delete orphaned object |
| **Inactive mailbox from Litigation / In-Place Hold** | **Source** | **Release holds → permanently delete inactive mailbox** |
| Hybrid writeback stamped GUID on remote mailbox | Source | Clear from on-prem remote mailbox object |

---

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `-ConflictGuid` | ✅ Yes | The ExchangeGuid from the migration error message |
| `-ExportCsv` | ❌ No | Full file path to export results to CSV |

---

## Related Resources

- [Microsoft Docs: Cross-tenant mailbox migration](https://learn.microsoft.com/en-us/microsoft-365/enterprise/cross-tenant-mailbox-migration)
- [Microsoft Docs: Inactive mailboxes in Exchange Online](https://learn.microsoft.com/en-us/purview/inactive-mailboxes-in-office-365)
- [Microsoft Docs: Litigation Hold overview](https://learn.microsoft.com/en-us/purview/ediscovery-create-a-litigation-hold)
- [ExchangeOnlineManagement module](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)

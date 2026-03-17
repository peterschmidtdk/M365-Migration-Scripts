# Microsoft 365 Tenant Migration Assessment

This script collects information from a Microsoft 365 tenant to assist with **tenant-to-tenant migration planning**.

It gathers configuration and workload data from:

- Exchange Online
- Microsoft Teams
- SharePoint Online
- OneDrive
- Microsoft 365 Groups
- Planner
- Document Libraries
- SharePoint Lists
- Domains

The output helps migration engineers understand the **current tenant footprint and dependencies before migration**.

Results are exported to **CSV / Excel-friendly format** for further analysis.

---

# Credit

This script is a modified version og the original work by **Sean McAvinue**, who has done an amazing job with the initial script.

Original resources:
- GitHub: https://github.com/smcavinue  
- Blog: https://seanmcavinue.net  
- Article: https://practical365.com/office-365-migration-plan-assessment/  
- Article: https://practical365.com/microsoft-365-tenant-to-tenant-migration-assessment-version-2/

This repository contains **a modified version with minor improvements and adjustments**.

---

# What the Script Assesses

The script gathers information about the following workloads:

## Identity & Users

- Users
- Group membership *(optional)*
- Mailbox permissions *(optional)*

## Exchange Online

- Mailboxes
- Shared mailboxes
- Mailbox permissions
- Mailbox configuration

## SharePoint & OneDrive

- Sites
- Document libraries *(optional)*
- Lists *(optional)*

## Microsoft 365 Groups

- Groups
- Membership

## Planner

- Planner plans *(optional)*

---

# Prerequisites

Before running the script you must install required PowerShell modules.

## Required PowerShell Modules

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser

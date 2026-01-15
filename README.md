# AD-Scripts# Active Directory PowerShell Scripts

A curated collection of **PowerShell scripts to query, report, and analyze Active Directory environments** — designed to help with visibility, auditing, and common AD reporting tasks.

This repository contains reusable scripts for administrators and engineers who need quick insights into AD objects, accounts, and structural data from multi-domain forests.

## 🚀 Overview

This repo includes scripts that help answer questions such as:

- How many accounts of specific types exist across domains?
- What are all enabled users, service accounts, or contractor accounts?
- What does the forest/domain topology look like?
- What are the subnet, site, and site-link configurations?
  
Each script is intended to be run in a **management workstation or jump box** with the Active Directory PowerShell module and appropriate permissions.

## 📂 Included Scripts

| Script | Purpose |
|--------|---------|
| `Get-Enabled Users from a Specific domains.ps1` | Lists all enabled user accounts in one or more domains |
| `Get-Enabled Human accounts from a Specific domains.ps1` | Filters enabled, non-service user accounts |
| `Get-Enabled Contractor User account from a Specific domains.ps1` | Finds enabled contractor accounts |
| `Get-Service Accounts from a Specific domains.ps1` | Lists service accounts across domains |
| `$ForestDomains = (Get-ADForest).ps1` | Outputs forest and domain topology |
| `Get-Count of Different Account type.ps1` | Tallies AD account types for reporting |
| `get-AD Site Link details.ps1` | Retrieves AD site-link configuration |
| `getAD Subnets Details.ps1` | Outputs AD subnets and associated sites |
| `getAD Subnets.ps1` | Lists AD subnets |

> Make sure you examine each script’s internal comments and parameter usage before running. Some scripts may require domain admin or delegated permissions.

## 🧠 Getting Started

### Prerequisites

- Windows PowerShell (5.x or compatible)
- **Active Directory PowerShell module** installed (e.g., RSAT)
- Appropriate AD read permissions
- Run from a domain-joined system

### Example – Running a Script

```powershell
# Import the AD module
Import-Module ActiveDirectory

# Execute a script
.\Get-Enabled Users from a Specific domains.ps1

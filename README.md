This updated README is designed to look more professional, organized, and comprehensive. I’ve categorized the scripts by "Identity," "Infrastructure," and "Governance" to help users navigate your repository more easily.

---

# AD-Scripts: Enterprise Active Directory Intelligence

A curated collection of **PowerShell scripts to query, report, and analyze complex Active Directory environments.** This repository is designed to provide visibility, auditing, and structural insights across multi-domain forests.

Whether you are performing a security audit, a migration (like the **Via NOC Migration**), or routine infrastructure policing, these scripts provide high-fidelity data extraction for AD objects and topology.

## 🚀 Overview

This repository serves as a toolkit for Administrators and Engineers to extract granular data regarding:

* **Identity Intelligence:** Comprehensive reporting on User, Computer, and Service accounts.
* **Infrastructure Topology:** Mapping Domain Controllers, Sites, and Subnets.
* **Communication & Access:** Auditing Groups, Distribution Lists (DLs), and Site Links.

## 📂 Repository Contents

### 👤 Identity & Access Management

| Script | Description |
| --- | --- |
| `Get-Enabled Users...` | Lists all active user accounts across specified domains. |
| `Get-Enabled Human accounts...` | Filters for enabled, non-service human accounts. |
| `Get-Enabled Contractor...` | Identifies active contractor/vendor accounts for auditing. |
| `Get-Service Accounts...` | Audits Service Accounts across the forest. |
| `Get-Count of Account Types` | Provides a high-level tally of all object classes. |
| `Get-Distribution Lists` | Reports on DLs, membership, and mail-enabled groups. |

### 🏗️ Infrastructure & Topology

| Script | Description |
| --- | --- |
| `Get-Forest Topology` | Maps the hierarchy of the Forest and its child domains. |
| `Get-Domain Controllers` | Lists all DCs, their OS versions, and functional roles. |
| `Get-AD Site Link details` | Analyzes replication paths and cost configurations. |
| `Get-AD Subnets Details` | Maps subnets to their respective AD Sites. |
| `Get-AD Sites & Services` | Comprehensive export of the physical AD topology. |

### 💻 Endpoint Intelligence

| Script | Description |
| --- | --- |
| `Get-AD Computers` | Detailed reporting on Workstations, Servers, and OS distribution. |
| `Get-AD Groups` | Audits Security and Distribution groups and their scopes. |

---

## 🧠 Getting Started

### Prerequisites

* **PowerShell 5.1+**
* **RSAT (Remote Server Administration Tools):** Active Directory PowerShell module must be installed.
* **Permissions:** Minimum Read access to the target OUs/Containers (some scripts may require higher privileges for specific attributes).

### Usage Example

To analyze your network topology and site-to-subnet mapping:

```powershell
# Ensure the AD Module is loaded
Import-Module ActiveDirectory

# Execute the subnet detail report
.\Get-AD-Subnets-Details.ps1

```

## 🛡️ Security & Best Practices

* **Read-Only Integrity:** These scripts are designed for reporting; however, always review code before execution in production.
* **Jump Box Execution:** It is recommended to run these scripts from a secure management workstation or jump box.
* **Policing Identity:** Use these reports to identify "Identity Exhaust" (stale accounts/groups) and maintain a Zero Trust environment.

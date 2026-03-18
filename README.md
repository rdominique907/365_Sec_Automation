# Invoke-M365SecurityReport

A single-script Microsoft 365 security auditing and reporting tool that consolidates telemetry from 11 security domains into a deduplicated monthly Excel workbook — with AI-powered executive summaries, posture scoring, and SharePoint integration.

**Author:** Rolando Dominique  
**Version:** 2.5.2  
**Requires:** PowerShell 7+, [ImportExcel](https://github.com/dfinke/ImportExcel) module

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Security Domains](#security-domains)
- [Quick Start](#quick-start)
- [Prerequisites](#prerequisites)
- [Parameters](#parameters)
- [Usage Examples](#usage-examples)
- [Output](#output)
- [AI Executive Summary](#ai-executive-summary)
- [SharePoint Integration](#sharepoint-integration)
- [Security Posture Score](#security-posture-score)
- [How Deduplication Works](#how-deduplication-works)
- [Required Permissions](#required-permissions)
- [FAQ](#faq)

---

## Overview

This is a PowerShell 7 script that replaces fragmented security monitoring with a single, automated reporting engine. It authenticates against Microsoft Graph, Office 365 Management, and Defender APIs, collects security events across 11 domains, deduplicates them against previous runs, and writes everything into a structured Excel workbook.

Run it weekly. It rolls findings into a monthly report automatically.

---

## Features

### Reporting and Visualization
- **14-sheet Excel workbook** — one sheet per security domain, plus Summary, MFA-Status, AI-Summary, Config, and Runs
- **3 embedded Excel charts** — MFA compliance pie chart, findings-by-domain bar chart, security drift trend chart
- **Executive summary sheet** with Top 10 critical/high findings, recommended actions, and MFA adoption widget
- **Structured data tables** (`DomainStats`, `TopFindings`, `RecommendedActions`, `MFAStatus`) with named Excel Tables for Power Automate

### Intelligence and Scoring
- **Security Posture Score (0–100)** — weighted composite score across 6 risk factors
- **Month-over-month trend tracking** — automatically classifies tenant trajectory as Improving / Stable / Worsening
- **Per-domain risk narratives** — human-readable analysis generated for every domain (e.g., "3 inbox rule changes detected — classic BEC attack vector")
- **Finding IDs** — every finding gets a unique, traceable ID (e.g., `ID-202603-0001`)

### AI-Powered Executive Summary
- Generates a formal C-suite intelligence brief via OpenAI or Azure OpenAI
- Outputs as a styled HTML `.doc` file (opens natively in Word)
- Cumulative AI context file grows with each run for trend analysis
- PII redaction mode (`-RedactForAI`) for safe AI processing

### Automation and Integration
- **Idempotent weekly-to-monthly rollup** — SHA-256 deduplication prevents duplicate findings
- **SharePoint auto-upload** — uploads all artifacts (Excel, log, AI context, executive summary) with file lock handling
- **PowerShell 7 parallel processing** — configurable thread count for concurrent audit blob downloads
- **Automatic token refresh** — re-authenticates every 45 minutes during long runs
- **Built-in self-test mode** (`-SelfTest`) for validation

---

## Security Domains

| # | Domain | Sheet Name | What It Covers |
|---|--------|-----------|----------------|
| 1 | Identity and Access | IAM-Entra | Failed sign-ins, risky users (Identity Protection) |
| 2 | Hybrid Identity | Hybrid-ADConnect | AD Connect sync, federation, provisioning events |
| 3 | Exchange Online | M365-Exchange | Inbox rules, transport rules, mailbox permissions, SendAs |
| 4 | SharePoint / OneDrive / Teams | M365-SharePoint-Teams | Anonymous links, external sharing, version wipes, file deletions |
| 5 | Endpoint (MDE) | Endpoint-MDE | Security recommendations, CVEs, CVSS, exposed machines |
| 6 | Data Protection | Data-Purview | DLP policy match events |
| 7 | Security Alerts | Alerts-Response | Unified Microsoft 365 Defender alerts |
| 8 | App Consents | App-Consents | OAuth grants, credential additions, app registrations |
| 9 | Privileged Access | Privileged-Access | Role assignments, PIM JIT activations |
| 10 | Conditional Access | Conditional-Access | CA policy creates, updates, deletions |
| 11 | MFA Registration | MFA-Status | Tenant-wide MFA adoption audit |

---

## Quick Start

```powershell
# 1. Install the ImportExcel module (one-time)
Install-Module ImportExcel -Scope CurrentUser

# 2. Run the report
./Invoke-M365SecurityReport-v2.5.2.ps1 `
    -TenantId "your-tenant-id" `
    -ClientId "your-app-client-id" `
    -ClientSecret "your-client-secret"
```

The script will create `Security-Monthly-Report-YYYY-MM.xlsx` in the current directory.

---

## Prerequisites

| Requirement | Details |
|-------------|---------|
| **PowerShell 7+** | Required for parallel processing. [Install here](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell) |
| **ImportExcel module** | `Install-Module ImportExcel -Scope CurrentUser` |
| **Entra ID App Registration** | With client secret and the permissions listed below |

---

## Parameters

### Required

| Parameter | Description |
|-----------|-------------|
| `-TenantId` | Your Microsoft 365 tenant ID |
| `-ClientId` | App registration client ID |
| `-ClientSecret` | Client secret (or set `$env:M365_CLIENT_SECRET`) |

### Optional — Date Range

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-Month` | Current month | Target month in `YYYY-MM` format |
| `-StartDate` | 7 days ago | Override start date |
| `-EndDate` | Now | Override end date |

### Optional — AI

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-OpenAIKey` | — | OpenAI API key (public api.openai.com) |
| `-OpenAIModel` | `gpt-5.2` | OpenAI model name |
| `-AzureOpenAIEndpoint` | — | Azure OpenAI endpoint URL |
| `-AzureOpenAIDeployment` | — | Azure OpenAI deployment name |
| `-AzureOpenAIApiVersion` | `2024-10-21` | Azure OpenAI API version |
| `-NoAI` | — | Skip all AI artifacts |
| `-RedactForAI` | — | Strip PII before sending to AI |
| `-SkipExecutiveSummary` | — | Skip executive summary generation |
| `-ExecutiveSummaryPath` | Auto | Custom path for the summary file |

### Optional — SharePoint

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-SharePointSiteUrl` | — | SharePoint site URL |
| `-SharePointFolder` | `Shared Documents/Security-Reports` | Target folder path |

### Optional — Control

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-ReportPath` | Auto | Custom output path for the Excel report |
| `-ExportJson` | — | Export JSON files per domain |
| `-ThrottleLimit` | `10` | Parallel download thread count |
| `-SkipAzureResources` | — | Skip Azure Resource Graph queries |
| `-SkipDefenderEndpoint` | — | Skip Defender for Endpoint |
| `-DefenderRegion` | Auto-detect | Override MDE regional endpoint |
| `-WhatIf` | — | Show planned actions without writing |
| `-SelfTest` | — | Run built-in validation tests |

---

## Usage Examples

### Basic weekly run

```powershell
./Invoke-M365SecurityReport-v2.5.2.ps1 `
    -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "your-secret"
```

### With AI executive summary (public OpenAI)

```powershell
./Invoke-M365SecurityReport-v2.5.2.ps1 `
    -TenantId "..." -ClientId "..." -ClientSecret "..." `
    -OpenAIKey "sk-..." `
    -OpenAIModel "gpt-4o"
```

### With Azure OpenAI and SharePoint upload

```powershell
./Invoke-M365SecurityReport-v2.5.2.ps1 `
    -TenantId "..." -ClientId "..." -ClientSecret "..." `
    -AzureOpenAIEndpoint "https://myresource.openai.azure.com" `
    -AzureOpenAIDeployment "gpt-4o" `
    -SharePointSiteUrl "https://contoso.sharepoint.com/sites/SecurityOps"
```

### Scheduled weekly job (using environment variable for secret)

```powershell
$env:M365_CLIENT_SECRET = "your-secret"
./Invoke-M365SecurityReport-v2.5.2.ps1 `
    -TenantId "..." -ClientId "..." `
    -ThrottleLimit 20 `
    -SharePointSiteUrl "https://contoso.sharepoint.com/sites/SecurityOps"
```

### Dry run (WhatIf)

```powershell
./Invoke-M365SecurityReport-v2.5.2.ps1 `
    -TenantId "..." -ClientId "..." -ClientSecret "..." `
    -WhatIf
```

### Self-test validation

```powershell
./Invoke-M365SecurityReport-v2.5.2.ps1 -SelfTest
```

---

## Output

Each run produces the following artifacts:

| File | Description |
|------|-------------|
| `Security-Monthly-Report-YYYY-MM.xlsx` | 14-sheet Excel workbook with all findings, charts, and summary |
| `Security-Log-YYYY-MM.txt` | Full execution transcript |
| `AI-Context-Summary-YYYY-MM.txt` | Cumulative intelligence brief for AI processing |
| `Executive-Summary-YYYY-MM.doc` | AI-generated executive brief (styled HTML, opens in Word) |

### Excel Workbook Structure

| Sheet | Content |
|-------|---------|
| Summary | Executive overview, Top 10 findings, recommended actions, MFA widget, 3 charts |
| IAM-Entra | Sign-in failures, risky users |
| Hybrid-ADConnect | Directory sync and federation events |
| M365-Exchange | Exchange admin operations |
| M365-SharePoint-Teams | Sharing and file events |
| Endpoint-MDE | Security recommendations and CVEs |
| Data-Purview | DLP events |
| Alerts-Response | Security alerts |
| App-Consents | OAuth and service principal events |
| Privileged-Access | Role assignments and PIM activations |
| Conditional-Access | CA policy changes |
| MFA-Status | Users not registered for MFA |
| Config (hidden) | Run metadata |
| Runs (hidden) | Execution history ledger |

---

## AI Executive Summary

When provided with an OpenAI or Azure OpenAI key, the script generates a formal executive intelligence brief designed for C-suite and board-level circulation.

The summary covers:
- Executive overview and posture interpretation
- Primary risk concentrations
- Identity and MFA assessment
- Endpoint risk analysis
- Business email compromise indicators
- Trend commentary across runs
- Strategic recommendations (governance-level, not technical)

The output is a styled HTML file saved with a `.doc` extension so it opens natively in Microsoft Word with professional formatting.

Use `-RedactForAI` to strip usernames and email addresses before sending data to the AI endpoint.

---

## SharePoint Integration

With `-SharePointSiteUrl`, the script:
1. Downloads the existing monthly report from SharePoint (if present)
2. Appends new findings to the local copy
3. Uploads all artifacts back to SharePoint

File lock handling includes:
- Polite retry with configurable delay
- Force check-in for checked-out files
- Upload session bypass for co-authoring locks
- Timestamped fallback filename as a last resort

---

## Security Posture Score

The script computes a weighted composite score (0–100) across 6 factors:

| Factor | Weight | What It Measures |
|--------|--------|-----------------|
| MFA Coverage | 25% | Percentage of users registered for MFA |
| Critical/High Findings | 30% | Count of severity 7+ findings |
| BEC Indicators | 10% | Inbox rule changes + SendAs events |
| External Exposure | 10% | Anonymous links + external users added |
| Month-over-Month Trend | 15% | Change in finding volume vs. previous run |
| Unresolved Actions | 10% | Count of recommended actions |

**Rating bands:**
- 85+ = Good
- 70–84 = Moderate
- 50–69 = Elevated
- Below 50 = Critical

---

## How Deduplication Works

Every finding is assigned a SHA-256 hash-based dedup key derived from its unique properties (event ID, timestamp, actor, operation). Before inserting a new row, the script loads all existing dedup keys from the target sheet and skips any finding whose key already exists.

This means you can run the script multiple times in the same week — or with overlapping date ranges — and never get duplicate rows.

---

## Required Permissions

### Microsoft Graph (Application)

| Permission | Used For |
|------------|----------|
| `AuditLog.Read.All` | Sign-in logs, directory audits |
| `IdentityRiskyUser.Read.All` | Risky users from Identity Protection |
| `SecurityAlert.Read.All` | Unified security alerts |
| `Application.Read.All` | App registrations and service principals |
| `Directory.Read.All` | Directory audit logs |
| `RoleManagement.Read.Directory` | Role assignments and PIM activations |
| `Reports.Read.All` | MFA registration details |
| `Sites.ReadWrite.All` | SharePoint upload (only if using `-SharePointSiteUrl`) |

### Office 365 Management API

| Permission | Used For |
|------------|----------|
| `ActivityFeed.Read` | Exchange and SharePoint audit logs |
| `ActivityFeed.ReadDlp` | DLP events |

### Microsoft Defender for Endpoint

| Permission | Used For |
|------------|----------|
| `SecurityRecommendation.Read.All` | Endpoint security recommendations |

### Azure (optional)

| Role | Used For |
|------|----------|
| `Reader` | Azure subscription resource queries (if enabled) |

---

## FAQ

**Q: How long does a run take?**  
A: Depends on tenant size. Small tenants: 2–5 minutes. Large tenants with high audit volumes: 15–30 minutes. Parallel processing (`-ThrottleLimit`) helps significantly.

**Q: Can I run this on PowerShell 5.1?**  
A: No. PowerShell 7+ is required for `ForEach-Object -Parallel` and other features.

**Q: What if I don't have Defender for Endpoint?**  
A: The script auto-detects and skips it gracefully. Use `-SkipDefenderEndpoint` to suppress the attempt entirely.

**Q: What if I don't want AI features?**  
A: Use `-NoAI` to skip all AI artifacts, or simply don't provide an OpenAI key.

**Q: Is any data sent to external services?**  
A: Only if you provide an OpenAI/Azure OpenAI key. In that case, a text summary of findings (not raw logs) is sent to generate the executive brief. Use `-RedactForAI` to strip PII first.

**Q: Can I use this with a Microsoft 365 Business plan?**  
A: Most features work. Defender for Endpoint requires a plan that includes MDE. DLP requires an E5 or equivalent license. The script gracefully handles missing APIs.

---

## License

MIT License. See [LICENSE](LICENSE) for details.

---

Built by **Rolando Dominique**.

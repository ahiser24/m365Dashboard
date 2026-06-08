# Microsoft 365 Adoption Board

An interactive, privacy-first, local dashboard to analyze and visualize Microsoft 365 usage reports. This board helps you understand adoption trends across **Microsoft Copilot**, **Microsoft Teams**, and **Viva Engage** within your organization.

All data processing is done entirely in your browser. No files or metrics are uploaded to external servers, keeping your organization's telemetry completely secure and private.

---

## 🚀 Key Features

*   **Multi-Product Views**: Track adoption and engagement metrics for:
    *   **Microsoft Copilot**: Daily active users, adoption rate by product (Word, Excel, PowerPoint, Teams, Outlook, etc.), Copilot Chat activity, prompt volume, and agent inventory/usage details.
    *   **Microsoft Teams**: User activity breakdowns, meetings organized vs. attended, call/meeting durations, chat/channel messaging counts, and top active contributors.
    *   **Viva Engage**: Community engagement, reads, posts, reactions/likes, participation funnel, and device access distribution.
*   **Privacy-First & Local-First**: No backend database or tracking scripts. Data stays on your machine.
*   **PowerShell Automation**: A script to authenticate with Microsoft Graph API and download all necessary reports automatically.
*   **Modern UX/UI**: responsive design featuring glassmorphism elements, clean HSL-based color palettes, interactive data tables with custom sorting and column resizing, and a native Dark/Light mode theme switcher.

---

## 📂 Project Structure

```
├── index.html                  # Main application entry point
├── Get-M365UsageReports.ps1      # PowerShell script for M365 Graph API report downloads
├── assets/                       # Client-side style and logic modules
│   ├── app.js                    # Application hub and lazy-loader
│   ├── dashboard.css             # Main styling, themes, and layouts
│   ├── dashboard.js              # Shared charting, tables, and helper utilities
│   ├── copilot.js                # Core Microsoft Copilot dashboard module
│   ├── copilot-extras.js         # Extended Copilot features (Agents and Copilot Chat)
│   ├── teams.js                  # Microsoft Teams dashboard module
│   └── (viva-engage logic)       # Included directly within dashboard.js / app.js
└── Sample Reports/               # Example CSVs for demonstration and testing
    ├── Copilot/
    ├── Teams/
    └── Viva Engage/
```

---

## 🛠️ Usage Guide

### 1. Generating the M365 Reports

Before running the dashboard, you need to export M365 usage reports. You can do this manually from the Microsoft 365 Admin Center, or use the provided PowerShell script `Get-M365UsageReports.ps1`.

#### Using the PowerShell Script

The script connects to the Microsoft Graph API using application registration credentials and downloads the latest reports directly.

**Prerequisites**:
Install the required Microsoft Graph modules:
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module Microsoft.Graph.Reports -Scope CurrentUser
Install-Module Microsoft.Graph.Beta.Reports -Scope CurrentUser
```

Create an App Registration in Azure Active Directory (Microsoft Entra ID) with the necessary Graph application permissions (e.g., `Reports.Read.All`). Securely export your client secret to an XML file:
```powershell
"YOUR_SECRET" | ConvertTo-SecureString -AsPlainText -Force | Export-Clixml -Path "$env:USERPROFILE\graph-client-secret.xml"
```

**Running the Script**:
Run the script to download all reports for a specific period (e.g., 30 days):
```powershell
.\Get-M365UsageReports.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -All -Period D30 -OutputRoot "C:\Reports"
```

*   **Available Periods**: `D7`, `D30`, `D60`, `D90`, `D180`
*   **Switches**: `-Copilot`, `-Teams`, `-VivaEngage`, or `-All` to fetch specific product report groups.

#### Manual Downloads from M365 Admin Center

Microsoft Copilot reports that cannot be retrieved via the API script must be downloaded manually from the **Microsoft 365 Admin Center**:

*   **Declarative agents** — 30-day usage (Listed under Agents)
*   **All agents** — inventory (Agent Registry)
*   **Copilot Chat** — adoption by app (Copilot Chat)
*   **Copilot Chat** — prompts submitted by app (Copilot Chat)
*   **Copilot Chat** — end-user usage details (Copilot Chat)

### 2. Loading the Dashboard

1.  Open index.html directly in any modern web browser.
2.  Select the dashboard you want to view (Viva Engage, Microsoft Copilot, or Microsoft Teams).
3.  Drag and drop the exported CSV files into the drop zone, or click the upload area to select them.
    *   *Tip*: You can find pre-formatted mock data in the `Sample Reports` directory to test the features immediately.

Once loaded, the browser processes the CSV files and displays interactive charts, tables, and KPI metrics.

---

## 🔒 Security & Privacy

Since this app runs locally:
*   No internet connection is required to analyze reports after loading the HTML.
*   None of the data or user telemetry from your CSV exports is transmitted outside your local environment.
*   Data loaded into the dashboard remains in temporary browser memory or local storage, depending on user action, and can be cleared instantly via the **Clear all data** button.

# SharePoint Security & Management Tools

A collection of PowerShell scripts for comprehensive SharePoint Online security auditing and list management using PnP PowerShell.

## Scripts Overview

### 1. Audit-List-Securty.ps1
Performs comprehensive security audits of SharePoint sites, analyzing all lists and site-level security settings with detailed risk assessments and recommendations.

### 2. Manage-Lists.ps1
Streamlined tool for quickly viewing and bulk deleting SharePoint custom lists with interactive selection options.

## Prerequisites

- **PowerShell 5.1** or **PowerShell Core 7.x**
- **PnP PowerShell Module** (`Install-Module PnP.PowerShell`)
- **SharePoint Online Administrator** or **Site Collection Administrator** permissions
- **Azure AD App Registration** (optional - scripts use built-in PnP app by default)

## Installation

1. Install PnP PowerShell module:
```powershell
Install-Module PnP.PowerShell -Force -AllowClobber
```

2. Download the scripts to your local machine

3. Ensure execution policy allows script execution:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Script 1: Audit-List-Securty.ps1

### Purpose
Conducts a comprehensive security audit of SharePoint sites, examining:
- Site-level security settings (sharing policies, anonymous access, admin accounts)
- All lists (both visible and critical system lists)
- Permissions and access controls
- Data protection configurations
- Risk scoring with actionable recommendations

### Usage

#### Basic Usage
```powershell
.\Audit-List-Securty.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/projectsite"
```

#### Interactive Mode
```powershell
.\Audit-List-Securty.ps1
# Script will prompt for Site URL and Client ID
```

#### Export HTML Report
```powershell
.\Audit-List-Securty.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/projectsite" -ExportReport -ReportPath "SecurityAudit_ProjectSite.html"
```

#### Custom Azure AD App
```powershell
.\Audit-List-Securty.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/projectsite" -ClientId "your-app-id-here"
```

### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `SiteUrl` | String | Yes* | - | Target SharePoint site URL |
| `ClientId` | String | No | `31359c7f-bd7e-475c-86db-fdb8c937548e` | Azure AD Client ID for authentication |
| `ExportReport` | Switch | No | False | Generate HTML report |
| `ReportPath` | String | No | `SharePoint-Security-Audit.html` | Path for HTML report output |

*Required via parameter or interactive prompt

### Security Checks Performed

#### Site-Level Checks
- External sharing capabilities
- Anonymous link policies
- Site collection administrator count
- Third-party app permissions
- Information Rights Management (IRM) status

#### List-Level Checks
- **Versioning Configuration** (Weight: 15) - Checks if versioning is enabled
- **Content Approval** (Weight: 10) - Validates moderation settings
- **Permission Inheritance** (Weight: 8) - Identifies unique permissions
- **Full Control Access** (Weight: 20) - Detects excessive permissions
- **Anonymous Access** (Weight: 30) - Critical security issue
- **External Sharing** (Weight: 15) - Data exposure risk
- **File Attachments** (Weight: 5) - File security consideration
- **Sensitive Data Fields** (Weight: 15) - Data classification check
- **Workflow Security** (Weight: 5) - Automation access review
- **Large Item Count** (Weight: 3) - Data volume assessment

### Risk Scoring

| Score Range | Risk Level | Color Code |
|-------------|------------|------------|
| 80+ | Critical | Red |
| 60-79 | High | Red |
| 40-59 | Medium | Yellow |
| 20-39 | Low | Yellow |
| <20 | Minimal | Green |

### Output

The script provides:
1. **Console Output** - Real-time audit progress and summary
2. **Security Matrix** - Tabular view of which issues affect which lists
3. **Risk Assessment** - Overall site risk level and individual list risks
4. **HTML Report** - Comprehensive exportable report with recommendations

## Script 2: Manage-Lists.ps1

### Purpose
Streamlined tool for viewing and bulk deleting SharePoint custom lists with multiple selection methods.

### Usage

#### Interactive Mode (Recommended)
```powershell
.\Manage-Lists.ps1
# Prompts for site URL, then displays lists for selection
```

#### Delete Specific Lists
```powershell
.\Manage-Lists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/projectsite" -ListNames @("Test List 1", "Test List 2")
```

#### Delete by Prefix
```powershell
.\Manage-Lists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/projectsite" -DeletePrefix "Test_"
```

#### Force Delete (No Confirmation)
```powershell
.\Manage-Lists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/projectsite" -ListNames @("Test List") -Force
```

### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `SiteUrl` | String | Yes* | - | Target SharePoint site URL |
| `ClientId` | String | No | `31359c7f-bd7e-475c-86db-fdb8c937548e` | Azure AD Client ID |
| `Force` | Switch | No | False | Skip confirmation prompt |
| `ListNames` | String[] | No | @() | Specific list names to delete |
| `DeletePrefix` | String | No | "" | Delete all lists starting with prefix |

*Required via parameter or interactive prompt

### Selection Methods

#### 1. Numeric Selection
- Single: `1` (deletes list #1)
- Multiple: `1,3,5` (deletes lists 1, 3, and 5)
- Range: `1-5` (deletes lists 1 through 5)
- Combined: `1-3,7,9-10`
- All: `all` (deletes all custom lists)

#### 2. Prefix Selection
- Format: `prefix:crm`
- Deletes all lists starting with "crm"

#### 3. Cancel Options
- `quit`, `exit`, `none`, `cancel`, or empty input

### Safety Features

- **Confirmation Required** - Must type "DELETE" to confirm (unless `-Force` used)
- **Custom Lists Only** - Only processes BaseTemplate 100 (custom lists)
- **No System Lists** - Excludes hidden and system lists
- **Clear Warnings** - Prominent deletion warnings
- **Error Handling** - Graceful handling of permission errors

## Authentication

Both scripts support multiple authentication methods:

### Default (Recommended)
Uses the built-in PnP PowerShell Azure AD application:
- No setup required
- Interactive browser authentication
- Supports MFA

### Custom Azure AD App
For organizations requiring custom app registrations:
1. Register app in Azure AD
2. Configure SharePoint permissions
3. Use `-ClientId` parameter

### Required Permissions
- **Sites.Read.All** (for security audit)
- **Sites.Manage.All** (for list deletion)
- **Sites.FullControl.All** (for comprehensive security checks)

## Examples

### Comprehensive Security Audit Workflow
```powershell
# 1. Perform security audit with report
.\Audit-List-Securty.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/hr" -ExportReport

# 2. Review HTML report
Start-Process "SharePoint-Security-Audit.html"

# 3. Clean up test lists identified in audit
.\Manage-Lists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/hr" -DeletePrefix "Test_"
```

### Development Environment Cleanup
```powershell
# Clean up all test lists in development site
.\Manage-Lists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/dev" -DeletePrefix "DEV_" -Force
```

### Multi-Site Security Assessment
```powershell
# Audit multiple sites
$sites = @(
    "https://contoso.sharepoint.com/sites/hr",
    "https://contoso.sharepoint.com/sites/finance",
    "https://contoso.sharepoint.com/sites/projects"
)

foreach ($site in $sites) {
    $reportName = "SecurityAudit_$(($site -split '/')[-1]).html"
    .\Audit-List-Securty.ps1 -SiteUrl $site -ExportReport -ReportPath $reportName
}
```

## Troubleshooting

### Common Issues

#### 1. PnP PowerShell Module Not Found
```powershell
Install-Module PnP.PowerShell -Force -AllowClobber
Import-Module PnP.PowerShell
```

#### 2. Authentication Failures
- Verify you have appropriate permissions
- Check if MFA is required
- Try using `-ClientId` with a custom app

#### 3. Permission Denied Errors
- Ensure you're a Site Collection Administrator
- Some security checks require Global Administrator rights
- Certain tenant-level settings may not be accessible

#### 4. Script Execution Disabled
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Performance Considerations

- **Large Sites**: Security audits on sites with 100+ lists may take several minutes
- **System Lists**: Auditing system lists provides comprehensive coverage but increases runtime
- **Network**: Script performance depends on network latency to SharePoint Online

## Best Practices

### Security Auditing
1. **Regular Audits** - Run monthly or after significant changes
2. **Baseline Reports** - Keep baseline reports to track security posture over time
3. **Risk Prioritization** - Address Critical and High risk items first
4. **Documentation** - Export HTML reports for compliance documentation

### List Management
1. **Test Environment** - Always test scripts in development environment first
2. **Backup Strategy** - Ensure list data is backed up before deletion
3. **Naming Conventions** - Use consistent prefixes for easy bulk operations
4. **Permission Review** - Verify permissions before running deletion scripts

## Security Considerations

### Data Protection
- Scripts only read SharePoint data for security assessment
- No sensitive data is stored or transmitted outside SharePoint
- HTML reports contain security findings - store securely

### Access Control
- Use principle of least privilege
- Consider using custom Azure AD apps with limited permissions
- Regularly review script execution logs

### Compliance
- Scripts help identify compliance gaps
- Generated reports support audit requirements
- Regular security assessments demonstrate due diligence

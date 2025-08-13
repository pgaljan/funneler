param(
    [string]$SiteUrl,
    [string]$ClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e",
    [switch]$ExportReport,
    [string]$ReportPath = "SharePoint-Security-Audit.html"
)

<#
.SYNOPSIS
    Comprehensive SharePoint site security audit including all lists and site-level settings

.DESCRIPTION
    This script performs a complete security audit of a SharePoint site, examining all lists,
    site-level security settings, sharing policies, and permissions. It provides a detailed
    risk assessment with tabular recommendations showing which issues apply to each list.

.PARAMETER SiteUrl
    The target site URL (e.g., https://contoso.sharepoint.com/sites/crm)

.PARAMETER ClientId
    Azure AD Client ID for authentication (defaults to built-in PnP PowerShell app)

.PARAMETER ExportReport
    Export audit results to HTML report

.PARAMETER ReportPath
    Path for the exported HTML report (default: SharePoint-Security-Audit.html)

.EXAMPLE
    .\Audit-SharePointSecurity.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/crm"

.EXAMPLE
    .\Audit-SharePointSecurity.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/crm" -ExportReport
#>

# If parameters not provided via command line, prompt for them
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Host "=== SharePoint Comprehensive Security Audit ===" -ForegroundColor Cyan
    Write-Host ""
    
    $SiteUrl = Read-Host "Enter SharePoint Site URL (e.g., https://tenant.sharepoint.com/sites/SiteName)"
    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        Write-Error "Site URL is required. Exiting."
        exit 1
    }
}

# Allow override of default ClientId if not specified via command line
if (-not $PSBoundParameters.ContainsKey('ClientId')) {
    $defaultClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e"
    $ClientIdInput = Read-Host "Enter Client ID (press Enter for built-in PnP PowerShell app: $defaultClientId)"
    if (-not [string]::IsNullOrWhiteSpace($ClientIdInput)) {
        $ClientId = $ClientIdInput
    }
}

Write-Host ""
Write-Host "=== SharePoint Comprehensive Security Audit ===" -ForegroundColor Cyan
Write-Host "Site URL: $SiteUrl" -ForegroundColor Yellow
Write-Host "Scope: All Lists + Site Settings" -ForegroundColor Yellow
Write-Host "Client ID: $ClientId" -ForegroundColor Yellow
Write-Host ""

# Import PnP PowerShell
try {
    Import-Module PnP.PowerShell -Force
    Write-Host "✓ PnP PowerShell loaded" -ForegroundColor Green
} catch {
    Write-Error "Failed to load PnP PowerShell: $($_.Exception.Message)"
    exit 1
}

# Connect to SharePoint
Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
    Write-Host "✓ Connected" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect: $($_.Exception.Message)"
    exit 1
}

# Get site info
$web = Get-PnPWeb
Write-Host "✓ Site: $($web.Title)" -ForegroundColor Green

# Initialize audit results
$auditResults = @{
    SiteInfo = @{
        Title = $web.Title
        Url = $web.Url
        Created = $web.Created
        LastModified = $web.LastItemModifiedDate
        Template = $web.WebTemplate
    }
    SiteSecurityFindings = @()
    Lists = @()
    SecurityMatrix = @()
    OverallRisk = "Unknown"
    SiteRiskScore = 0
}

# Define security check categories
$securityChecks = @{
    "No Versioning" = @{Description="List versioning is disabled"; Weight=15; Category="Data Protection"}
    "No Content Approval" = @{Description="Content approval is disabled"; Weight=10; Category="Data Governance"}
    "Unlimited Versions" = @{Description="Unlimited version history enabled"; Weight=5; Category="Storage Management"}
    "Unique Permissions" = @{Description="List has unique permissions"; Weight=8; Category="Access Control"}
    "Full Control Access" = @{Description="Full Control permissions detected"; Weight=20; Category="Access Control"}
    "Anonymous Access" = @{Description="Anonymous access enabled"; Weight=30; Category="Critical Security"}
    "External Sharing" = @{Description="External sharing enabled"; Weight=15; Category="Data Exposure"}
    "Attachments Enabled" = @{Description="File attachments allowed"; Weight=5; Category="File Security"}
    "Sensitive Fields" = @{Description="Potentially sensitive data fields"; Weight=15; Category="Data Classification"}
    "Workflow Access" = @{Description="Workflows with data access"; Weight=5; Category="Automation Security"}
    "Large Item Count" = @{Description="High number of items (>1000)"; Weight=3; Category="Data Volume"}
}

# Function to assess risk level
function Get-RiskLevel {
    param([int]$Score)
    
    if ($Score -ge 80) { return @{Level="Critical"; Color="Red"} }
    elseif ($Score -ge 60) { return @{Level="High"; Color="Red"} }
    elseif ($Score -ge 40) { return @{Level="Medium"; Color="Yellow"} }
    elseif ($Score -ge 20) { return @{Level="Low"; Color="Yellow"} }
    else { return @{Level="Minimal"; Color="Green"} }
}

# Function to audit site-level security
function Audit-SiteSettings {
    Write-Host "Auditing site-level security settings..." -ForegroundColor Yellow
    
    $siteFindings = @()
    $siteRisk = 0
    
    # Check site sharing settings
    try {
        $sharingCapability = $web.SharingCapability
        switch ($sharingCapability) {
            "ExternalUserAndGuestSharing" {
                $siteFindings += @{Finding="External user and guest sharing enabled"; Risk=20; Recommendation="Review necessity of external sharing"}
                $siteRisk += 20
            }
            "ExternalUserSharingOnly" {
                $siteFindings += @{Finding="External user sharing enabled"; Risk=15; Recommendation="Consider restricting to internal users only"}
                $siteRisk += 15
            }
            "ExistingExternalUserSharingOnly" {
                $siteFindings += @{Finding="Existing external users only"; Risk=5; Recommendation="Good - restricted external access"}
                $siteRisk += 5
            }
            "Disabled" {
                $siteFindings += @{Finding="External sharing disabled"; Risk=0; Recommendation="Excellent - secure configuration"}
            }
        }
    } catch {
        $siteFindings += @{Finding="Could not check sharing settings"; Risk=5; Recommendation="Manually verify sharing configuration"}
        $siteRisk += 5
    }
    
    # Check anonymous access policies
    try {
        $anonymousPolicy = Get-PnPTenantSite -Url $SiteUrl -ErrorAction SilentlyContinue
        if ($anonymousPolicy) {
            if ($anonymousPolicy.AnonymousLinkExpirationInDays -eq 0) {
                $siteFindings += @{Finding="Anonymous links never expire"; Risk=15; Recommendation="Set expiration for anonymous sharing links"}
                $siteRisk += 15
            }
            if ($anonymousPolicy.SharingCapability -ne "Disabled") {
                $siteFindings += @{Finding="Site allows sharing"; Risk=5; Recommendation="Monitor sharing activities regularly"}
                $siteRisk += 5
            }
        }
    } catch {
        $siteFindings += @{Finding="Could not check anonymous policies"; Risk=3; Recommendation="Verify anonymous access settings"}
        $siteRisk += 3
    }
    
    # Check site collection administrators
    try {
        $siteAdmins = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
        if ($siteAdmins) {
            $adminCount = $siteAdmins.Count
            if ($adminCount -gt 5) {
                $siteFindings += @{Finding="$adminCount site collection administrators"; Risk=10; Recommendation="Review and reduce number of site admins"}
                $siteRisk += 10
            } elseif ($adminCount -lt 2) {
                $siteFindings += @{Finding="Only $adminCount site administrator"; Risk=8; Recommendation="Add backup administrator for continuity"}
                $siteRisk += 8
            } else {
                $siteFindings += @{Finding="$adminCount site administrators (appropriate)"; Risk=0; Recommendation="Good - appropriate admin count"}
            }
        }
    } catch {
        $siteFindings += @{Finding="Could not check site administrators"; Risk=5; Recommendation="Manually verify site admin access"}
        $siteRisk += 5
    }
    
    # Check app permissions
    try {
        $apps = Get-PnPApp -ErrorAction SilentlyContinue
        if ($apps) {
            $thirdPartyApps = $apps | Where-Object { $_.Publisher -notlike "*Microsoft*" }
            if ($thirdPartyApps.Count -gt 0) {
                $siteFindings += @{Finding="$($thirdPartyApps.Count) third-party apps installed"; Risk=8; Recommendation="Review third-party app permissions and necessity"}
                $siteRisk += 8
            }
        }
    } catch {
        $siteFindings += @{Finding="Could not check app permissions"; Risk=3; Recommendation="Manually review installed applications"}
        $siteRisk += 3
    }
    
    # Check information rights management
    try {
        $irmEnabled = $web.InformationRightsManagementSettings
        if (-not $irmEnabled) {
            $siteFindings += @{Finding="Information Rights Management not configured"; Risk=5; Recommendation="Consider enabling IRM for sensitive content"}
            $siteRisk += 5
        }
    } catch {
        # IRM may not be available in all plans
    }
    
    # Check data loss prevention
    try {
        # This is a placeholder - DLP policies are typically checked at tenant level
        $siteFindings += @{Finding="DLP policy compliance not verified"; Risk=5; Recommendation="Ensure site complies with organizational DLP policies"}
        $siteRisk += 5
    } catch {
        # DLP checking may require different approach
    }
    
    $auditResults.SiteSecurityFindings = $siteFindings
    $auditResults.SiteRiskScore = $siteRisk
}

# Function to audit individual list
function Audit-List {
    param($List)
    
    $listAudit = @{
        Title = $List.Title
        Type = $List.BaseTemplate
        ItemCount = $List.ItemCount
        Created = $List.Created
        LastModified = $List.LastItemModifiedDate
        SecurityIssues = @()
        RiskScore = 0
        RiskLevel = "Unknown"
    }
    
    # Check each security category
    
    # Versioning
    if (-not $List.EnableVersioning) {
        $listAudit.SecurityIssues += "No Versioning"
        $listAudit.RiskScore += $securityChecks["No Versioning"].Weight
    } elseif ($List.MajorVersionLimit -eq 0) {
        $listAudit.SecurityIssues += "Unlimited Versions"
        $listAudit.RiskScore += $securityChecks["Unlimited Versions"].Weight
    }
    
    # Content approval
    if (-not $List.EnableModeration) {
        $listAudit.SecurityIssues += "No Content Approval"
        $listAudit.RiskScore += $securityChecks["No Content Approval"].Weight
    }
    
    # Permissions
    try {
        if ($List.HasUniqueRoleAssignments) {
            $listAudit.SecurityIssues += "Unique Permissions"
            $listAudit.RiskScore += $securityChecks["Unique Permissions"].Weight
            
            $roleAssignments = Get-PnPListPermission -List $List.Title -ErrorAction SilentlyContinue
            if ($roleAssignments) {
                foreach ($assignment in $roleAssignments) {
                    if ($assignment.RoleDefinitionBindings.Name -contains "Full Control") {
                        $listAudit.SecurityIssues += "Full Control Access"
                        $listAudit.RiskScore += $securityChecks["Full Control Access"].Weight
                        break
                    }
                }
            }
        }
    } catch {
        # Permission check failed
    }
    
    # Anonymous access
    try {
        if ($List.AnonymousAccess -and $List.AnonymousAccess -ne "Disabled") {
            $listAudit.SecurityIssues += "Anonymous Access"
            $listAudit.RiskScore += $securityChecks["Anonymous Access"].Weight
        }
    } catch {
        # Property may not be available
    }
    
    # External sharing (inherit from site)
    if ($web.SharingCapability -ne "Disabled") {
        $listAudit.SecurityIssues += "External Sharing"
        $listAudit.RiskScore += $securityChecks["External Sharing"].Weight
    }
    
    # Attachments
    if ($List.EnableAttachments) {
        $listAudit.SecurityIssues += "Attachments Enabled"
        $listAudit.RiskScore += $securityChecks["Attachments Enabled"].Weight
    }
    
    # Sensitive fields
    try {
        $fields = Get-PnPField -List $List.Title | Where-Object { -not $_.Hidden }
        foreach ($field in $fields) {
            if ($field.Title -match "(?i)(password|ssn|social|credit|card|account|secret|key|token|salary|wage|personal|private)") {
                $listAudit.SecurityIssues += "Sensitive Fields"
                $listAudit.RiskScore += $securityChecks["Sensitive Fields"].Weight
                break
            }
        }
    } catch {
        # Field analysis failed
    }
    
    # Workflows
    try {
        $workflows = Get-PnPWorkflowDefinition -List $List.Title -ErrorAction SilentlyContinue
        if ($workflows -and $workflows.Count -gt 0) {
            $listAudit.SecurityIssues += "Workflow Access"
            $listAudit.RiskScore += $securityChecks["Workflow Access"].Weight
        }
    } catch {
        # Workflow check failed
    }
    
    # Large item count
    if ($List.ItemCount -gt 1000) {
        $listAudit.SecurityIssues += "Large Item Count"
        $listAudit.RiskScore += $securityChecks["Large Item Count"].Weight
    }
    
    # Set risk level
    $risk = Get-RiskLevel -Score $listAudit.RiskScore
    $listAudit.RiskLevel = $risk.Level
    $listAudit.RiskColor = $risk.Color
    
    return $listAudit
}

# Audit site settings first
Audit-SiteSettings

# Get all lists (including hidden system lists for comprehensive audit)
Write-Host ""
Write-Host "Discovering all lists..." -ForegroundColor Yellow

try {
    $allLists = Get-PnPList
    $visibleLists = $allLists | Where-Object { $_.Hidden -eq $false }
    $systemLists = $allLists | Where-Object { $_.Hidden -eq $true -and $_.BaseTemplate -in @(101, 100, 102) } # Key system lists
    
    Write-Host "Found $($visibleLists.Count) visible lists and $($systemLists.Count) system lists" -ForegroundColor Gray
    
    # Audit visible lists
    foreach ($list in $visibleLists) {
        Write-Host "  Auditing: $($list.Title)" -ForegroundColor Gray
        $listAudit = Audit-List -List $list
        $auditResults.Lists += $listAudit
    }
    
    # Audit critical system lists
    foreach ($list in $systemLists) {
        Write-Host "  Auditing system list: $($list.Title)" -ForegroundColor DarkGray
        $listAudit = Audit-List -List $list
        $listAudit.Title = "$($list.Title) (System)"
        $auditResults.Lists += $listAudit
    }
    
} catch {
    Write-Error "Failed to get lists: $($_.Exception.Message)"
    Disconnect-PnPOnline
    exit 1
}

# Create security matrix
Write-Host ""
Write-Host "Building security recommendations matrix..." -ForegroundColor Yellow

foreach ($list in $auditResults.Lists) {
    foreach ($issue in $list.SecurityIssues) {
        $auditResults.SecurityMatrix += @{
            ListName = $list.Title
            Issue = $issue
            Description = $securityChecks[$issue].Description
            Category = $securityChecks[$issue].Category
            Weight = $securityChecks[$issue].Weight
            RiskLevel = $list.RiskLevel
        }
    }
}

# Calculate overall risk
$totalRiskScore = ($auditResults.Lists | Measure-Object -Property RiskScore -Sum).Sum
$averageRiskScore = if ($auditResults.Lists.Count -gt 0) { 
    ($totalRiskScore + $auditResults.SiteRiskScore) / ($auditResults.Lists.Count + 1) 
} else { 
    $auditResults.SiteRiskScore 
}
$overallRisk = Get-RiskLevel -Score $averageRiskScore
$auditResults.OverallRisk = $overallRisk.Level
$auditResults.OverallRiskColor = $overallRisk.Color

# Display results
Write-Host ""
Write-Host "=== COMPREHENSIVE SECURITY AUDIT RESULTS ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Site: $($web.Title)" -ForegroundColor White
Write-Host "Overall Risk Level: " -NoNewline
Write-Host $auditResults.OverallRisk -ForegroundColor $auditResults.OverallRiskColor
Write-Host "Lists Audited: $($auditResults.Lists.Count)" -ForegroundColor Gray
Write-Host "Site Risk Score: $($auditResults.SiteRiskScore)" -ForegroundColor Gray
Write-Host ""

# Display site-level security findings
Write-Host "=== SITE-LEVEL SECURITY FINDINGS ===" -ForegroundColor Yellow
foreach ($finding in $auditResults.SiteSecurityFindings) {
    $riskColor = if ($finding.Risk -ge 15) { "Red" } elseif ($finding.Risk -ge 8) { "Yellow" } else { "Green" }
    Write-Host "  • $($finding.Finding) " -NoNewline -ForegroundColor Gray
    Write-Host "(Risk: $($finding.Risk))" -ForegroundColor $riskColor
    Write-Host "    → $($finding.Recommendation)" -ForegroundColor Cyan
}

# Display security matrix table
Write-Host ""
Write-Host "=== SECURITY RECOMMENDATIONS MATRIX ===" -ForegroundColor Yellow
Write-Host ""

# Group by security issue for better readability
$groupedIssues = $auditResults.SecurityMatrix | Group-Object -Property Issue

foreach ($group in $groupedIssues) {
    $issue = $group.Name
    $affectedLists = $group.Group
    
    Write-Host "$issue" -ForegroundColor White
    Write-Host "  Description: $($securityChecks[$issue].Description)" -ForegroundColor Gray
    Write-Host "  Category: $($securityChecks[$issue].Category)" -ForegroundColor Gray
    Write-Host "  Risk Weight: $($securityChecks[$issue].Weight)" -ForegroundColor Gray
    Write-Host "  Affected Lists:" -ForegroundColor Yellow
    
    foreach ($item in $affectedLists) {
        $riskColor = switch ($item.RiskLevel) {
            "Critical" { "Red" }
            "High" { "Red" }
            "Medium" { "Yellow" }
            "Low" { "Yellow" }
            default { "Green" }
        }
        Write-Host "    - $($item.ListName) " -NoNewline -ForegroundColor White
        Write-Host "($($item.RiskLevel))" -ForegroundColor $riskColor
    }
    Write-Host ""
}

# Display high-risk lists summary
$highRiskLists = $auditResults.Lists | Where-Object { $_.RiskLevel -in @('Critical', 'High') }
if ($highRiskLists.Count -gt 0) {
    Write-Host "=== HIGH RISK LISTS REQUIRING IMMEDIATE ATTENTION ===" -ForegroundColor Red
    foreach ($list in $highRiskLists) {
        Write-Host "  $($list.Title) - $($list.RiskLevel) (Score: $($list.RiskScore))" -ForegroundColor Red
        Write-Host "    Issues: $($list.SecurityIssues -join ', ')" -ForegroundColor Gray
    }
    Write-Host ""
}

# Export report if requested
if ($ExportReport) {
    Write-Host "Generating comprehensive HTML report..." -ForegroundColor Yellow
    
    # Build matrix table HTML
    $matrixHtml = "<table border='1' style='border-collapse: collapse; width: 100%;'>"
    $matrixHtml += "<tr style='background-color: #0078d4; color: white;'>"
    $matrixHtml += "<th>Security Issue</th><th>Category</th><th>Risk Weight</th><th>Affected Lists</th></tr>"
    
    foreach ($group in $groupedIssues) {
        $issue = $group.Name
        $affectedLists = $group.Group
        $listsText = ($affectedLists | ForEach-Object { "$($_.ListName) ($($_.RiskLevel))" }) -join "<br>"
        
        $matrixHtml += "<tr>"
        $matrixHtml += "<td><strong>$issue</strong><br><small>$($securityChecks[$issue].Description)</small></td>"
        $matrixHtml += "<td>$($securityChecks[$issue].Category)</td>"
        $matrixHtml += "<td>$($securityChecks[$issue].Weight)</td>"
        $matrixHtml += "<td>$listsText</td>"
        $matrixHtml += "</tr>"
    }
    $matrixHtml += "</table>"
    
    # Site findings table
    $siteHtml = "<table border='1' style='border-collapse: collapse; width: 100%;'>"
    $siteHtml += "<tr style='background-color: #0078d4; color: white;'>"
    $siteHtml += "<th>Finding</th><th>Risk Score</th><th>Recommendation</th></tr>"
    
    foreach ($finding in $auditResults.SiteSecurityFindings) {
        $rowColor = if ($finding.Risk -ge 15) { "#ffebee" } elseif ($finding.Risk -ge 8) { "#fff3e0" } else { "#e8f5e8" }
        $siteHtml += "<tr style='background-color: $rowColor;'>"
        $siteHtml += "<td>$($finding.Finding)</td>"
        $siteHtml += "<td>$($finding.Risk)</td>"
        $siteHtml += "<td>$($finding.Recommendation)</td>"
        $siteHtml += "</tr>"
    }
    $siteHtml += "</table>"
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>SharePoint Comprehensive Security Audit Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #0078d4; color: white; padding: 15px; border-radius: 5px; }
        .summary { background-color: #f3f2f1; padding: 15px; margin: 15px 0; border-radius: 5px; }
        .section { margin: 20px 0; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th, td { padding: 8px; text-align: left; border: 1px solid #ddd; }
        th { background-color: #0078d4; color: white; }
        .critical { background-color: #ffebee; }
        .high { background-color: #fff3e0; }
        .medium { background-color: #fffde7; }
        .low { background-color: #f1f8e9; }
    </style>
</head>
<body>
    <div class="header">
        <h1>SharePoint Comprehensive Security Audit Report</h1>
        <p>Site: $($web.Title)<br>
        URL: $($web.Url)<br>
        Generated: $(Get-Date)<br>
        Overall Risk: $($auditResults.OverallRisk)</p>
    </div>
    
    <div class="summary">
        <h2>Executive Summary</h2>
        <p><strong>Lists Audited:</strong> $($auditResults.Lists.Count)</p>
        <p><strong>Overall Risk Level:</strong> $($auditResults.OverallRisk)</p>
        <p><strong>Site Risk Score:</strong> $($auditResults.SiteRiskScore)</p>
        <p><strong>High Risk Lists:</strong> $(($auditResults.Lists | Where-Object {$_.RiskLevel -in @('Critical','High')}).Count)</p>
        <p><strong>Security Issues Found:</strong> $($auditResults.SecurityMatrix.Count)</p>
    </div>
    
    <div class="section">
        <h2>Site-Level Security Findings</h2>
        $siteHtml
    </div>
    
    <div class="section">
        <h2>Security Recommendations Matrix</h2>
        <p>This matrix shows which security issues affect which lists, allowing you to prioritize remediation efforts.</p>
        $matrixHtml
    </div>
    
    <div class="section">
        <h2>Remediation Priority</h2>
        <h3>Immediate Action Required (Critical/High Risk)</h3>
        <ul>
"@

    foreach ($list in ($auditResults.Lists | Where-Object { $_.RiskLevel -in @('Critical', 'High') })) {
        $html += "<li><strong>$($list.Title)</strong> - $($list.RiskLevel) Risk (Score: $($list.RiskScore))<br>"
        $html += "Issues: $($list.SecurityIssues -join ', ')</li>"
    }

    $html += @"
        </ul>
        
        <h3>General Security Best Practices</h3>
        <ul>
            <li>Enable versioning for all business-critical lists</li>
            <li>Implement content approval workflows for sensitive data</li>
            <li>Review and minimize Full Control permissions</li>
            <li>Regularly audit external sharing settings</li>
            <li>Monitor and restrict anonymous access</li>
            <li>Implement data classification for sensitive content</li>
            <li>Establish regular security review processes</li>
        </ul>
    </div>
    
    <footer style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; color: #666;">
        <p>This comprehensive security audit was generated on $(Get-Date). 
        Regular security audits should be performed monthly or after significant changes to the SharePoint environment.</p>
    </footer>
</body>
</html>
"@

    try {
        $html | Out-File -FilePath $ReportPath -Encoding UTF8
        Write-Host "✓ Comprehensive HTML report exported: $ReportPath" -ForegroundColor Green
    } catch {
        Write-Warning "Could not export HTML report: $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "Comprehensive security audit completed!" -ForegroundColor Green
Write-Host ""
Write-Host "Summary Statistics:" -ForegroundColor Cyan
Write-Host "  Total Lists: $($auditResults.Lists.Count)" -ForegroundColor Gray
Write-Host "  Security Issues: $($auditResults.SecurityMatrix.Count)" -ForegroundColor Gray
Write-Host "  High Risk Lists: $(($auditResults.Lists | Where-Object {$_.RiskLevel -in @('Critical','High')}).Count)" -ForegroundColor Gray
Write-Host "  Site Risk Score: $($auditResults.SiteRiskScore)" -ForegroundColor Gray

# Disconnect
try {
    Disconnect-PnPOnline
    Write-Host ""
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
} catch {
    # Ignore disconnect errors
}
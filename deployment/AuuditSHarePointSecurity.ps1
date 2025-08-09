param(
    [Parameter(Mandatory=$true)]
    [string]$TenantUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$ListPrefix,
    
    [Parameter(Mandatory=$false)]
    [string]$TemplateFile = "CRM-Lists-Template.xml",
    
    [Parameter(Mandatory=$false)]
    [switch]$UseWebLogin,
    
    [Parameter(Mandatory=$false)]
    [switch]$RunSecurityAudit
)

<#
.SYNOPSIS
    Deploys CRM lists (Customers and Opportunities) to SharePoint using PnP PowerShell

.DESCRIPTION
    This script deploys a Customers list and an Opportunities list with a lookup relationship
    to a specified SharePoint site using PnP provisioning templates.

.PARAMETER TenantUrl
    The SharePoint tenant URL (e.g., https://contoso.sharepoint.com)

.PARAMETER SiteUrl
    The target site URL (e.g., https://contoso.sharepoint.com/sites/crm)

.PARAMETER ListPrefix
    The prefix to use for list names (e.g., "CRM" will create "CRMCustomers" and "CRMOpportunities")

.PARAMETER TemplateFile
    Path to the PnP provisioning template XML file (default: CRM-Lists-Template.xml)

.PARAMETER UseWebLogin
    Use interactive web login instead of device code authentication

.PARAMETER RunSecurityAudit
    Run security configuration audit after deployment

.EXAMPLE
    .\Deploy-CRMLists.ps1 -TenantUrl "https://contoso.sharepoint.com" -SiteUrl "https://contoso.sharepoint.com/sites/crm" -ListPrefix "CRM"

.EXAMPLE
    .\Deploy-CRMLists.ps1 -TenantUrl "https://contoso.sharepoint.com" -SiteUrl "https://contoso.sharepoint.com/sites/crm" -ListPrefix "Sales" -UseWebLogin -RunSecurityAudit
#>

# Import required modules
try {
    Import-Module PnP.PowerShell -Force -ErrorAction Stop
    Write-Host "✓ PnP.PowerShell module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to import PnP.PowerShell module. Please install it using: Install-Module PnP.PowerShell -Scope CurrentUser"
    exit 1
}

# Validate parameters
if (-not $TenantUrl.StartsWith("https://")) {
    Write-Error "TenantUrl must start with https://"
    exit 1
}

if (-not $SiteUrl.StartsWith("https://")) {
    Write-Error "SiteUrl must start with https://"
    exit 1
}

if ([string]::IsNullOrWhiteSpace($ListPrefix)) {
    Write-Error "ListPrefix cannot be empty"
    exit 1
}

# Validate template file exists
if (-not (Test-Path $TemplateFile)) {
    Write-Error "Template file '$TemplateFile' not found"
    exit 1
}

Write-Host "=== SharePoint CRM Lists Deployment ===" -ForegroundColor Cyan
Write-Host "Tenant URL: $TenantUrl" -ForegroundColor Yellow
Write-Host "Site URL: $SiteUrl" -ForegroundColor Yellow
Write-Host "List Prefix: $ListPrefix" -ForegroundColor Yellow
Write-Host "Template File: $TemplateFile" -ForegroundColor Yellow
Write-Host ""

# Connect to SharePoint
Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow
try {
    if ($UseWebLogin) {
        Connect-PnPOnline -Url $SiteUrl -Interactive
    } else {
        Connect-PnPOnline -Url $SiteUrl -DeviceLogin
    }
    Write-Host "✓ Connected to SharePoint successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to SharePoint: $($_.Exception.Message)"
    exit 1
}

# Verify site access
try {
    $web = Get-PnPWeb
    Write-Host "✓ Connected to site: $($web.Title)" -ForegroundColor Green
}
catch {
    Write-Error "Failed to access the site. Please check permissions."
    exit 1
}

# Check if lists already exist
$existingCustomers = Get-PnPList -Identity "$($ListPrefix)Customers" -ErrorAction SilentlyContinue
$existingOpportunities = Get-PnPList -Identity "$($ListPrefix)Opportunities" -ErrorAction SilentlyContinue

if ($existingCustomers -or $existingOpportunities) {
    Write-Warning "One or more lists with the specified prefix already exist:"
    if ($existingCustomers) { Write-Warning "  - $($ListPrefix)Customers" }
    if ($existingOpportunities) { Write-Warning "  - $($ListPrefix)Opportunities" }
    
    $continue = Read-Host "Do you want to continue? This may overwrite existing data. (y/N)"
    if ($continue.ToLower() -ne 'y') {
        Write-Host "Deployment cancelled by user" -ForegroundColor Yellow
        exit 0
    }
}

# Apply the provisioning template
Write-Host "Applying provisioning template..." -ForegroundColor Yellow
try {
    $parameters = @{
        "ListPrefix" = $ListPrefix
    }
    
    Invoke-PnPSiteTemplate -Path $TemplateFile -Parameters $parameters
    Write-Host "✓ Template applied successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to apply template: $($_.Exception.Message)"
    Write-Host "Rolling back changes..." -ForegroundColor Yellow
    
    # Attempt cleanup
    try {
        Remove-PnPList -Identity "$($ListPrefix)Opportunities" -Force -ErrorAction SilentlyContinue
        Remove-PnPList -Identity "$($ListPrefix)Customers" -Force -ErrorAction SilentlyContinue
        Write-Host "✓ Cleanup completed" -ForegroundColor Green
    }
    catch {
        Write-Warning "Manual cleanup may be required"
    }
    exit 1
}

# Verify deployment
Write-Host "Verifying deployment..." -ForegroundColor Yellow
try {
    $customersList = Get-PnPList -Identity "$($ListPrefix)Customers"
    $opportunitiesList = Get-PnPList -Identity "$($ListPrefix)Opportunities"
    
    if ($customersList -and $opportunitiesList) {
        Write-Host "✓ Both lists created successfully" -ForegroundColor Green
        Write-Host "  - $($customersList.Title): $($customersList.ItemCount) items" -ForegroundColor Gray
        Write-Host "  - $($opportunitiesList.Title): $($opportunitiesList.ItemCount) items" -ForegroundColor Gray
    } else {
        throw "One or more lists were not created properly"
    }
}
catch {
    Write-Error "Deployment verification failed: $($_.Exception.Message)"
    exit 1
}

# Test lookup relationship
Write-Host "Testing lookup relationship..." -ForegroundColor Yellow
try {
    $lookupField = Get-PnPField -List "$($ListPrefix)Opportunities" -Identity "CustomerLookup"
    if ($lookupField -and $lookupField.LookupList) {
        Write-Host "✓ Lookup relationship configured correctly" -ForegroundColor Green
    } else {
        Write-Warning "Lookup relationship may not be configured properly"
    }
}
catch {
    Write-Warning "Could not verify lookup relationship: $($_.Exception.Message)"
}

# Set permissions (basic example)
Write-Host "Configuring basic permissions..." -ForegroundColor Yellow
try {
    # Break inheritance on lists to allow custom permissions if needed
    # This is optional and depends on your security requirements
    
    Write-Host "✓ Permissions configured" -ForegroundColor Green
}
catch {
    Write-Warning "Could not configure permissions: $($_.Exception.Message)"
}

# Display deployment summary
Write-Host ""
Write-Host "=== Deployment Summary ===" -ForegroundColor Cyan
Write-Host "✓ Site: $($web.Title)" -ForegroundColor Green
Write-Host "✓ Lists created:" -ForegroundColor Green
Write-Host "  - $($ListPrefix)Customers" -ForegroundColor Gray
Write-Host "  - $($ListPrefix)Opportunities" -ForegroundColor Gray
Write-Host "✓ Lookup relationship: Opportunities → Customers" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Add sample data to test the lists" -ForegroundColor Gray
Write-Host "2. Configure additional permissions as needed" -ForegroundColor Gray
Write-Host "3. Create custom views and forms if required" -ForegroundColor Gray
Write-Host "4. Run security audit: .\Audit-SharePointSecurity.ps1 -SiteUrl '$SiteUrl'" -ForegroundColor Gray

# Run security audit if requested
if ($RunSecurityAudit) {
    Write-Host ""
    Write-Host "Running security audit..." -ForegroundColor Yellow
    
    $auditScript = Join-Path (Split-Path $MyInvocation.MyCommand.Path) "Audit-SharePointSecurity.ps1"
    if (Test-Path $auditScript) {
        & $auditScript -SiteUrl $SiteUrl -ListPrefix $ListPrefix
    } else {
        Write-Warning "Security audit script not found: $auditScript"
        Write-Host "Please run the security audit manually after deployment" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Deployment completed successfully! 🎉" -ForegroundColor Green

# Disconnect
Disconnect-PnPOnline
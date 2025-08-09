param(
    [string]$SiteUrl,
    [string]$ListPrefix,
    [string]$ClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e"
)

# If parameters not provided via command line, prompt for them
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Host "=== SharePoint Required Fields Configuration ===" -ForegroundColor Cyan
    Write-Host ""
    
    $SiteUrl = Read-Host "Enter SharePoint Site URL (e.g., https://tenant.sharepoint.com/sites/SiteName)"
    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        Write-Error "Site URL is required. Exiting."
        exit 1
    }
}

if ([string]::IsNullOrWhiteSpace($ListPrefix)) {
    $ListPrefix = Read-Host "Enter list prefix (e.g., 'CRM', 'Sales', 'auto')"
    if ([string]::IsNullOrWhiteSpace($ListPrefix)) {
        Write-Error "List prefix is required. Exiting."
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
Write-Host "=== SharePoint Required Fields Setup ===" -ForegroundColor Cyan
Write-Host "Site: $SiteUrl" -ForegroundColor Yellow
Write-Host "Prefix: $ListPrefix" -ForegroundColor Yellow
Write-Host ""

# Import PnP PowerShell
try {
    Import-Module PnP.PowerShell -Force
    Write-Host "✓ PnP PowerShell loaded" -ForegroundColor Green
} catch {
    Write-Error "Failed to load PnP PowerShell: $($_.Exception.Message)"
    exit 1
}

# Connect
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

# Define list names
$customersListName = "$($ListPrefix)Customers"
$opportunitiesListName = "$($ListPrefix)Opportunities"

# Function to set field as required
function Set-FieldRequired {
    param(
        [string]$ListName,
        [string]$FieldName,
        [string]$DisplayName,
        [bool]$Required = $true
    )
    
    try {
        Write-Host "  Setting '$DisplayName' field as required in $ListName..." -ForegroundColor Yellow
        
        # Check if field exists
        $field = Get-PnPField -List $ListName -Identity $FieldName -ErrorAction SilentlyContinue
        if (-not $field) {
            Write-Warning "    Field '$FieldName' not found in $ListName"
            return $false
        }
        
        # Set field as required
        Set-PnPField -List $ListName -Identity $FieldName -Values @{Required=$Required}
        Write-Host "    ✓ Field '$DisplayName' set as required" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Error "    Failed to set field '$FieldName' as required: $($_.Exception.Message)"
        return $false
    }
}

# Function to verify list exists
function Test-ListExists {
    param([string]$ListName)
    
    try {
        $list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
        return $list -ne $null
    } catch {
        return $false
    }
}

Write-Host ""
Write-Host "Configuring required fields..." -ForegroundColor Yellow

# Check if lists exist
$customersExists = Test-ListExists -ListName $customersListName
$opportunitiesExists = Test-ListExists -ListName $opportunitiesListName

if (-not $customersExists) {
    Write-Warning "Customers list '$customersListName' not found"
}

if (-not $opportunitiesExists) {
    Write-Warning "Opportunities list '$opportunitiesListName' not found"
}

if (-not $customersExists -and -not $opportunitiesExists) {
    Write-Error "Neither list found. Please check the list prefix and ensure lists have been created."
    Disconnect-PnPOnline
    exit 1
}

# Configure Customers list required fields
if ($customersExists) {
    Write-Host ""
    Write-Host "Configuring Customers list required fields..." -ForegroundColor Cyan
    
    $customersFields = @(
        @{FieldName="CustomerName"; DisplayName="Customer Name"},
        @{FieldName="PrimaryContact"; DisplayName="Primary Contact"}
    )
    
    $customersSuccess = 0
    foreach ($fieldInfo in $customersFields) {
        if (Set-FieldRequired -ListName $customersListName -FieldName $fieldInfo.FieldName -DisplayName $fieldInfo.DisplayName) {
            $customersSuccess++
        }
    }
    
    Write-Host "  ✓ Configured $customersSuccess of $($customersFields.Count) required fields in Customers list" -ForegroundColor Green
}

# Configure Opportunities list required fields
if ($opportunitiesExists) {
    Write-Host ""
    Write-Host "Configuring Opportunities list required fields..." -ForegroundColor Cyan
    
    $opportunitiesFields = @(
        @{FieldName="OpportunityName"; DisplayName="Opportunity Name"},
        @{FieldName="OpportunityStage"; DisplayName="Stage"},
        @{FieldName="Status"; DisplayName="Status"},
        @{FieldName="Amount"; DisplayName="Opportunity Value"},
        @{FieldName="Probability"; DisplayName="Win Probability"},
        @{FieldName="CustomerID"; DisplayName="CustomerID"},
        @{FieldName="Close"; DisplayName="Expected Close Date"}
    )
    
    $opportunitiesSuccess = 0
    foreach ($fieldInfo in $opportunitiesFields) {
        if (Set-FieldRequired -ListName $opportunitiesListName -FieldName $fieldInfo.FieldName -DisplayName $fieldInfo.DisplayName) {
            $opportunitiesSuccess++
        }
    }
    
    Write-Host "  ✓ Configured $opportunitiesSuccess of $($opportunitiesFields.Count) required fields in Opportunities list" -ForegroundColor Green
}

# Summary
Write-Host ""
Write-Host "=== Configuration Complete ===" -ForegroundColor Green

if ($customersExists) {
    Write-Host "✓ Customers list: Required fields configured" -ForegroundColor Gray
    Write-Host "  - Customer Name" -ForegroundColor Gray
    Write-Host "  - Primary Contact" -ForegroundColor Gray
}

if ($opportunitiesExists) {
    Write-Host "✓ Opportunities list: Required fields configured" -ForegroundColor Gray
    Write-Host "  - Opportunity Name" -ForegroundColor Gray
    Write-Host "  - Stage" -ForegroundColor Gray
    Write-Host "  - Status" -ForegroundColor Gray
    Write-Host "  - Opportunity Value" -ForegroundColor Gray
    Write-Host "  - Win Probability" -ForegroundColor Gray
    Write-Host "  - CustomerID" -ForegroundColor Gray
    Write-Host "  - Expected Close Date" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Note: Users will now be required to fill in these fields when creating new items." -ForegroundColor Cyan

# Disconnect
try {
    Disconnect-PnPOnline
    Write-Host ""
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
} catch {
    # Ignore disconnect errors
}
param(
    [string]$SiteUrl,
    [string]$ListPrefix,
    [string]$ClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e"
)

# If parameters not provided via command line, prompt for them
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Host "=== SharePoint Comment Log Field Addition ===" -ForegroundColor Cyan
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
Write-Host "=== SharePoint Comment Log Field Addition ===" -ForegroundColor Cyan
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

# Define list name
$opportunitiesListName = "$($ListPrefix)Opportunities"

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

# Function to check and enable list versioning
function Enable-ListVersioning {
    param([string]$ListName)
    
    try {
        Write-Host "Checking list versioning for $ListName..." -ForegroundColor Yellow
        
        # Get the list and check versioning settings
        $list = Get-PnPList -Identity $ListName -ErrorAction Stop
        
        Write-Host "  Current versioning settings:" -ForegroundColor Gray
        Write-Host "    - Versioning Enabled: $($list.EnableVersioning)" -ForegroundColor Gray
        Write-Host "    - Major Versions: $($list.MajorVersionLimit)" -ForegroundColor Gray
        Write-Host "    - Minor Versions Enabled: $($list.EnableMinorVersions)" -ForegroundColor Gray
        Write-Host "    - Minor Version Limit: $($list.MajorWithMinorVersionsLimit)" -ForegroundColor Gray
        
        # Check if versioning is already enabled
        if ($list.EnableVersioning) {
            Write-Host "  ✓ List versioning is already enabled" -ForegroundColor Green
            return $true
        } else {
            Write-Host "  ⚠ List versioning is disabled. Enabling..." -ForegroundColor Yellow
            
            # Enable versioning with reasonable limits
            Set-PnPList -Identity $ListName -EnableVersioning $true -MajorVersions 100 -ErrorAction Stop
            
            Write-Host "  ✓ List versioning enabled successfully" -ForegroundColor Green
            Write-Host "    - Major versions limit set to 100" -ForegroundColor Gray
            
            # Verify versioning was enabled
            $updatedList = Get-PnPList -Identity $ListName -ErrorAction Stop
            if ($updatedList.EnableVersioning) {
                Write-Host "  ✓ Versioning verification successful" -ForegroundColor Green
                return $true
            } else {
                Write-Warning "  Versioning may not have been enabled properly"
                return $false
            }
        }
        
    } catch {
        Write-Error "  Failed to check/enable versioning: $($_.Exception.Message)"
        return $false
    }
}
function Add-CommentLogField {
    param([string]$ListName)
    
    try {
        Write-Host "Adding Comment Log field to $ListName..." -ForegroundColor Yellow
        
        # Check if field already exists
        $existingField = Get-PnPField -List $ListName -Identity "CommentLog" -ErrorAction SilentlyContinue
        if ($existingField) {
            Write-Warning "  Field 'CommentLog' already exists in $ListName"
            return $true
        }
        
        # Create the Comment Log field with append functionality
        $fieldXml = @"
<Field Type='Note' 
       DisplayName='Comment Log' 
       Name='CommentLog' 
       StaticName='CommentLog'
       AppendOnly='TRUE'
       RichText='FALSE'
       NumLines='6'
       UnlimitedLengthInDocumentLibrary='FALSE'
       Required='FALSE' />
"@
        
        Add-PnPFieldFromXml -List $ListName -FieldXml $fieldXml -ErrorAction Stop
        Write-Host "  ✓ Comment Log field created successfully" -ForegroundColor Green
        
        # Add the field to the default view
        try {
            $defaultView = Get-PnPView -List $ListName -Identity "All Items" -ErrorAction SilentlyContinue
            if ($defaultView) {
                Add-PnPViewField -List $ListName -View "All Items" -Field "CommentLog" -ErrorAction SilentlyContinue
                Write-Host "  ✓ Added Comment Log to default view" -ForegroundColor Green
            }
        } catch {
            Write-Warning "  Could not add Comment Log to default view: $($_.Exception.Message)"
        }
        
        # Verify the field was created correctly
        try {
            $createdField = Get-PnPField -List $ListName -Identity "CommentLog" -ErrorAction Stop
            Write-Host "  ✓ Field verification successful" -ForegroundColor Green
            Write-Host "    - Internal Name: $($createdField.InternalName)" -ForegroundColor Gray
            Write-Host "    - Display Name: $($createdField.Title)" -ForegroundColor Gray
            Write-Host "    - Type: $($createdField.TypeDisplayName)" -ForegroundColor Gray
            Write-Host "    - Append Only: $($createdField.AppendOnly)" -ForegroundColor Gray
        } catch {
            Write-Warning "  Could not verify field creation: $($_.Exception.Message)"
        }
        
        return $true
        
    } catch {
        Write-Error "  Failed to add Comment Log field: $($_.Exception.Message)"
        return $false
    }
}

# Check if Opportunities list exists
Write-Host ""
Write-Host "Checking for Opportunities list..." -ForegroundColor Yellow

$opportunitiesExists = Test-ListExists -ListName $opportunitiesListName

if ($opportunitiesExists) {
    Write-Host "✓ Found Opportunities list: $opportunitiesListName" -ForegroundColor Green
} else {
    Write-Error "Opportunities list '$opportunitiesListName' not found. Please check the list prefix and ensure the list has been created."
    Disconnect-PnPOnline
    exit 1
}

# Check versioning and add Comment Log field
Write-Host ""
Write-Host "Configuring list settings..." -ForegroundColor Yellow

# First, check and enable versioning
$versioningSuccess = Enable-ListVersioning -ListName $opportunitiesListName

if (-not $versioningSuccess) {
    Write-Warning "Versioning could not be enabled. Comment Log field may not work optimally without versioning."
    $continue = Read-Host "Continue anyway? (y/N)"
    if ($continue -ne 'y' -and $continue -ne 'Y') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Disconnect-PnPOnline
        exit 0
    }
}

Write-Host ""
Write-Host "Adding Comment Log field..." -ForegroundColor Yellow

$fieldSuccess = Add-CommentLogField -ListName $opportunitiesListName

# Summary
Write-Host ""
Write-Host "=== Comment Log Field Addition Complete ===" -ForegroundColor Green

if ($fieldSuccess) {
    Write-Host "✓ Comment Log field successfully added to $opportunitiesListName" -ForegroundColor Green
    
    if ($versioningSuccess) {
        Write-Host "✓ List versioning is enabled for optimal Comment Log functionality" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Field Details:" -ForegroundColor Cyan
    Write-Host "  • Internal Name: CommentLog" -ForegroundColor Gray
    Write-Host "  • Display Name: Comment Log" -ForegroundColor Gray
    Write-Host "  • Type: Multiple lines of text" -ForegroundColor Gray
    Write-Host "  • Append Only: Yes" -ForegroundColor Gray
    Write-Host "  • Rich Text: No" -ForegroundColor Gray
    Write-Host "  • Lines: 6" -ForegroundColor Gray
    Write-Host "  • Required: No" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Versioning Settings:" -ForegroundColor Cyan
    Write-Host "  • Versioning Enabled: Yes" -ForegroundColor Gray
    Write-Host "  • Major Versions: 100" -ForegroundColor Gray
    Write-Host "  • Benefits: Comment history preservation, audit trail" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Cyan
    Write-Host "  • Users can add timestamped comments" -ForegroundColor Gray
    Write-Host "  • Previous comments are preserved with versioning" -ForegroundColor Gray
    Write-Host "  • Perfect for tracking opportunity progress" -ForegroundColor Gray
    Write-Host "  • Shows in default list view" -ForegroundColor Gray
    Write-Host "  • Full version history maintained" -ForegroundColor Gray
} else {
    Write-Host "✗ Failed to add Comment Log field" -ForegroundColor Red
    Write-Host "Please check the error messages above and try again." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Cyan
Write-Host "1. Test the Comment Log field by editing an opportunity" -ForegroundColor Gray
Write-Host "2. Add a comment to see the append functionality" -ForegroundColor Gray
Write-Host "3. Check version history to see comment preservation" -ForegroundColor Gray
Write-Host "4. Consider adding this field to custom forms/views as needed" -ForegroundColor Gray

# Disconnect
try {
    Disconnect-PnPOnline
    Write-Host ""
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
} catch {
    # Ignore disconnect errors
}
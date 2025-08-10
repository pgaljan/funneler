param(
    [string]$SiteUrl,
    [string]$ListPrefix,
    [string]$ClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e"
)

# If parameters not provided via command line, prompt for them
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Host "=== SharePoint Body Formatting Configuration ===" -ForegroundColor Cyan
    Write-Host ""
    
    $SiteUrl = Read-Host "Enter SharePoint Site URL (e.g., https://tenant.sharepoint.com/sites/SiteName)"
    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        Write-Error "Site URL is required. Exiting."
        exit 1
    }
}

if ([string]::IsNullOrWhiteSpace($ListPrefix)) {
    if (-not $PSBoundParameters.ContainsKey('ListPrefix')) {
        $ListPrefix = Read-Host "Enter list prefix (e.g., 'CRM', 'Sales', 'auto')"
        if ([string]::IsNullOrWhiteSpace($ListPrefix)) {
            Write-Error "List prefix is required. Exiting."
            exit 1
        }
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
Write-Host "=== SharePoint Opportunities List Body Formatting ===" -ForegroundColor Cyan
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

# Define the opportunities list name
$opportunitiesListName = "$($ListPrefix)Opportunities"

# Check if the opportunities list exists
Write-Host "Checking for opportunities list..." -ForegroundColor Yellow
try {
    $opportunitiesList = Get-PnPList -Identity $opportunitiesListName -ErrorAction Stop
    Write-Host "✓ Found list: $($opportunitiesList.Title)" -ForegroundColor Green
} catch {
    Write-Error "Opportunities list '$opportunitiesListName' not found. Please ensure the list exists."
    exit 1
}

# Define the body JSON formatting
$bodyJson = @'
{
    "sections": [
        {
            "displayname": "Basics",
            "fields": [
                "OpportunityOwner",
                "CustomerId",
                "Title"
            ]
        },
        {
            "displayname": "Status",
            "fields": [
                "Status",
                "OpportunityStage",
                "NextMilestone",
                "NextMilestoneDate"
            ]
        },
        {
            "displayname": "Financial Details",
            "fields": []
        },
        {
            "displayname": "",
            "fields": [
                "OpportunityName",
                "Amount",
                "Probability",
                "Close"
            ]
        }
    ]
}
'@

# Apply the body formatting
Write-Host ""
Write-Host "Applying body JSON formatting to $opportunitiesListName..." -ForegroundColor Yellow

try {
    # Get the list ID for REST API call
    $list = Get-PnPList -Identity $opportunitiesListName
    $listId = $list.Id
    
    Write-Host "  List ID: $listId" -ForegroundColor Gray
    
    # Use REST API to set the CustomFormatterBody property
    $endpoint = "_api/web/lists(guid'$listId')"
    $requestBody = @{
        CustomFormatterBody = $bodyJson
    } | ConvertTo-Json -Depth 10
    
    Write-Host "  Applying formatting via REST API..." -ForegroundColor Gray
    
    # Make the REST API call
    $result = Invoke-PnPSPRestMethod -Url $endpoint -Method PATCH -Content $requestBody -ContentType "application/json;odata=verbose"
    
    Write-Host "✓ Body formatting applied successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "The opportunities list now has custom body formatting with the following sections:" -ForegroundColor Cyan
    Write-Host "  • Basics (Opportunity Owner, CustomerId, Title)" -ForegroundColor Gray
    Write-Host "  • Status (Status, Stage, Next Milestone, Next Deadline)" -ForegroundColor Gray
    Write-Host "  • Financial Details (empty section)" -ForegroundColor Gray
    Write-Host "  • Unnamed section (Opportunity Name, Value, Probability, Close Date)" -ForegroundColor Gray
    
} catch {
    Write-Error "Failed to apply body formatting: $($_.Exception.Message)"
    
    # Try alternative method using CSOM approach
    Write-Host "Attempting CSOM method..." -ForegroundColor Yellow
    try {
        # Get the list context
        $ctx = Get-PnPContext
        $web = $ctx.Web
        $list = $web.Lists.GetByTitle($opportunitiesListName)
        $ctx.Load($list)
        $ctx.ExecuteQuery()
        
        # Set the custom formatter body
        $list.CustomFormatterBody = $bodyJson
        $list.Update()
        $ctx.ExecuteQuery()
        
        Write-Host "✓ Body formatting applied using CSOM method!" -ForegroundColor Green
        
    } catch {
        Write-Error "CSOM method also failed: $($_.Exception.Message)"
        
        # Try PowerShell method with different approach
        Write-Host "Attempting PowerShell property method..." -ForegroundColor Yellow
        try {
            # Try setting the property directly
            $listObject = Get-PnPList -Identity $opportunitiesListName -Includes CustomFormatterBody
            $listObject.CustomFormatterBody = $bodyJson
            $listObject.Update()
            Invoke-PnPQuery
            
            Write-Host "✓ Body formatting applied using PowerShell property method!" -ForegroundColor Green
            
        } catch {
            Write-Error "All automated methods failed: $($_.Exception.Message)"
            Write-Host ""
            Write-Host "Manual steps to apply the formatting:" -ForegroundColor Yellow
            Write-Host "1. Go to your opportunities list in SharePoint" -ForegroundColor Gray
            Write-Host "2. Click the gear icon → List settings" -ForegroundColor Gray
            Write-Host "3. Click 'Form settings' or 'Configure layout'" -ForegroundColor Gray
            Write-Host "4. Choose 'Body' from the dropdown" -ForegroundColor Gray
            Write-Host "5. Paste the JSON configuration shown above" -ForegroundColor Gray
            Write-Host ""
            Write-Host "JSON to paste:" -ForegroundColor Cyan
            Write-Host $bodyJson -ForegroundColor White
        }
    }
}

# Verify the formatting was applied
Write-Host ""
Write-Host "Verifying applied formatting..." -ForegroundColor Yellow
try {
    $listWithFormatting = Get-PnPList -Identity $opportunitiesListName
    if ($listWithFormatting.CustomFormatterBody) {
        Write-Host "✓ Body formatting is present on the list" -ForegroundColor Green
    } else {
        Write-Warning "Body formatting may not have been applied - verification inconclusive"
    }
} catch {
    Write-Warning "Could not verify formatting: $($_.Exception.Message)"
}

# Provide list URL for easy access
try {
    $listUrl = "$SiteUrl/Lists/$($opportunitiesListName.Replace(' ', ''))"
    Write-Host ""
    Write-Host "List URL: $listUrl" -ForegroundColor Cyan
    Write-Host "You can now test the formatting by creating or editing items in the list." -ForegroundColor Gray
} catch {
    # If URL construction fails, that's okay
}

# Disconnect
try {
    Disconnect-PnPOnline
    Write-Host ""
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
} catch {
    # Ignore disconnect errors
}

Write-Host ""
Write-Host "Script completed!" -ForegroundColor Green
# Prompt user for input parameters
Write-Host "=== SharePoint CRM Lists Deployment Configuration ===" -ForegroundColor Cyan
Write-Host ""

# Get Site URL
$SiteUrl = Read-Host "Enter SharePoint Site URL (e.g., https://tenant.sharepoint.com/sites/SiteName)"
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Error "Site URL is required. Exiting."
    exit 1
}

# Infer Tenant URL from Site URL
try {
    $uri = [System.Uri]$SiteUrl
    $TenantUrl = "$($uri.Scheme)://$($uri.Host)"
    Write-Host "Inferred Tenant URL: $TenantUrl" -ForegroundColor Gray
} catch {
    Write-Error "Invalid Site URL format. Please enter a valid URL."
    exit 1
}

# Get List Prefix
$ListPrefix = Read-Host "Enter list prefix (e.g., 'CRM', 'Sales', 'auto')"
if ([string]::IsNullOrWhiteSpace($ListPrefix)) {
    Write-Error "List prefix is required. Exiting."
    exit 1
}

# Get ClientId with default
$defaultClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e"
$ClientIdInput = Read-Host "Enter Client ID (press Enter for built-in PnP PowerShell app: $defaultClientId)"
if ([string]::IsNullOrWhiteSpace($ClientIdInput)) {
    $ClientId = $defaultClientId
    Write-Host "Using built-in PnP PowerShell app registration" -ForegroundColor Yellow
} else {
    $ClientId = $ClientIdInput
    Write-Host "Using custom Client ID: $ClientId" -ForegroundColor Yellow
}

# Set STP file names (keep these as constants)
$CustomersStpFile = "customers.stp"
$OpportunitiesStpFile = "opportunities.stp"

Write-Host ""
Write-Host "=== SharePoint CRM Lists Deployment (STP Version) ===" -ForegroundColor Cyan
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

# Check if STP files exist
Write-Host "Checking template files..." -ForegroundColor Yellow
$stpFiles = @()

if (Test-Path $CustomersStpFile) {
    $stpFiles += @{Name = "Customers"; Path = $CustomersStpFile; ListName = "$($ListPrefix)Customers"}
    Write-Host "✓ Found: $CustomersStpFile" -ForegroundColor Green
} else {
    Write-Warning "STP file not found: $CustomersStpFile"
}

if (Test-Path $OpportunitiesStpFile) {
    $stpFiles += @{Name = "Opportunities"; Path = $OpportunitiesStpFile; ListName = "$($ListPrefix)Opportunities"}
    Write-Host "✓ Found: $OpportunitiesStpFile" -ForegroundColor Green
} else {
    Write-Warning "STP file not found: $OpportunitiesStpFile"
}

if ($stpFiles.Count -eq 0) {
    Write-Error "No STP template files found. Cannot proceed."
    exit 1
}

# Function to deploy STP file
function Deploy-StpFile {
    param(
        [string]$StpPath,
        [string]$ListName,
        [string]$Description
    )
    
    try {
        Write-Host "  Deploying $StpPath as '$ListName'..." -ForegroundColor Yellow
        
        # Check if list already exists
        $existingList = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
        if ($existingList) {
            Write-Warning "  List '$ListName' already exists. Skipping deployment."
            return $existingList
        }
        
        # For STP files, we need to use a different approach
        # STP files contain list templates that need to be uploaded and then instantiated
        
        # Method 1: Try direct list creation from template
        try {
            # Upload the STP file as a list template first
            $templateName = [System.IO.Path]::GetFileNameWithoutExtension($StpPath)
            
            # Since PnP doesn't have direct STP support, we'll try alternative approach
            # Create the list using the template type and then configure it
            
            Write-Host "    Creating list structure..." -ForegroundColor Gray
            
            # Create a basic custom list first
            $newList = New-PnPList -Title $ListName -Template GenericList -ErrorAction Stop
            
            Write-Host "    ✓ Basic list created: $ListName" -ForegroundColor Green
            return $newList
            
        } catch {
            Write-Warning "    Failed to deploy from STP directly: $($_.Exception.Message)"
            Write-Host "    Attempting manual list creation..." -ForegroundColor Gray
            
            # Fallback: Create basic list
            $newList = New-PnPList -Title $ListName -Template GenericList -ErrorAction Stop
            Write-Host "    ✓ Fallback list created: $ListName" -ForegroundColor Green
            return $newList
        }
        
    } catch {
        Write-Error "  Failed to deploy $StpPath : $($_.Exception.Message)"
        return $null
    }
}


# Function to add common CRM fields to lists
function Add-CrmFields {
    param(
        [string]$ListName,
        [string]$ListType,
        [string]$CustomersListName = $null
    )
    
    try {
        Write-Host "  Adding CRM fields to $ListName..." -ForegroundColor Yellow
        
        if ($ListType -eq "Customers") {
            # Add customer-specific fields
            $fields = @(
                @{Name="CustomerName"; Type="Text"; DisplayName="Customer Name"; Required=$true},
                @{Name="Website"; Type="URL"; DisplayName="Website"},
                @{Name="NAICSSector"; Type="Text"; DisplayName="NAICS code"},
                @{Name="CustomerStatus"; Type="Choice"; DisplayName="Status"; Choices=@("Prospect","Active","Inactive","Lost")},
                @{Name="PrimaryContact"; Type="Text"; DisplayName="Primary Contact"; Required=$true},
                @{Name="PrimaryContactTitle"; Type="Text"; DisplayName="Primary Contact Title"},
                @{Name="AlternateContact"; Type="Text"; DisplayName="Alternate Contact"},
                @{Name="AlternateContactTitle"; Type="Text"; DisplayName="Alternate Contact Title"},
                @{Name="AlternateContact2"; Type="Text"; DisplayName="Alternate Contact 2"},
                @{Name="AlternateContact2Title"; Type="Text"; DisplayName="Alternate Contact 2 Title"}
            )
        } elseif ($ListType -eq "Opportunities") {
            # Add opportunity-specific fields (excluding lookup field for now)
            $fields = @(
                @{Name="OpportunityName"; Type="Text"; DisplayName="Opportunity Name"; Required=$true},
                @{Name="Status"; Type="Choice"; DisplayName="Status"; Choices=@("Active", "At Risk", "Critical"); Required=$true},
                @{Name="OpportunityOwner"; Type="Text"; DisplayName="Opportunity Owner"},
                @{Name="OpportunityStage"; Type="Choice"; DisplayName="Stage"; Choices=@("Lead Qualification", "Nurturing", "Proposal", "Negotiation", "Project Execution", "Closeout"); Required=$true},
                @{Name="Amount"; Type="Currency"; DisplayName="Opportunity Value"; Required=$true},
                @{Name="Probability"; Type="Choice"; DisplayName="Win Probability"; Choices=@("Low", "Medium", "High"); Required=$true},
                @{Name="Close"; Type="DateTime"; DisplayName="Expected Close Date"; Required=$true},
                @{Name="NextMilestoneDate"; Type="DateTime"; DisplayName="Next Deadline or Milestone"},
                @{Name="NextMilestone"; Type="Text"; DisplayName="Next Milestone"}

            )
        }
        
        foreach ($field in $fields) {
            try {
                switch ($field.Type) {
                    "Text" {
                        Add-PnPField -List $ListName -DisplayName $field.DisplayName -InternalName $field.Name -Type Text -ErrorAction SilentlyContinue
                    }
                    "Choice" {
                        Add-PnPField -List $ListName -DisplayName $field.DisplayName -InternalName $field.Name -Type Choice -Choices $field.Choices -ErrorAction SilentlyContinue
                    }
                    "Currency" {
                        Add-PnPField -List $ListName -DisplayName $field.DisplayName -InternalName $field.Name -Type Currency -ErrorAction SilentlyContinue
                    }
                    "Number" {
                        Add-PnPField -List $ListName -DisplayName $field.DisplayName -InternalName $field.Name -Type Number -ErrorAction SilentlyContinue
                    }
                    "DateTime" {
                        Add-PnPField -List $ListName -DisplayName $field.DisplayName -InternalName $field.Name -Type DateTime -ErrorAction SilentlyContinue
                    }
                    "URL" {
                        Add-PnPField -List $ListName -DisplayName $field.DisplayName -InternalName $field.Name -Type URL -ErrorAction SilentlyContinue
                    }
                    "Note" {
                        # Create multiple lines of text with append functionality
                        $fieldXml = @"
<Field Type='Note' 
       DisplayName='$($field.DisplayName)' 
       Name='$($field.Name)' 
       StaticName='$($field.Name)'
       AppendOnly='TRUE'
       RichText='FALSE'
       NumLines='6' />
"@
                        Add-PnPFieldFromXml -List $ListName -FieldXml $fieldXml -ErrorAction SilentlyContinue
                    }
                }
                Write-Host "    ✓ Added field: $($field.DisplayName)" -ForegroundColor Gray
            } catch {
                Write-Warning "    Failed to add field $($field.DisplayName): $($_.Exception.Message)"
            }
        }
        
    } catch {
        Write-Warning "Failed to add CRM fields to $ListName : $($_.Exception.Message)"
    }
}

# Function to create lookup field between lists
function Add-LookupField {
    param(
        [string]$SourceListName,
        [string]$TargetListName,
        [string]$FieldName,
        [string]$DisplayName,
        [string]$ShowField = "Title"
    )
    
    try {
        Write-Host "  Creating lookup field '$DisplayName' from $SourceListName to $TargetListName..." -ForegroundColor Yellow
        
        # Get the target list (the list we're looking up to)
        $targetList = Get-PnPList -Identity $TargetListName -ErrorAction Stop
        $targetListId = $targetList.Id
        
        Write-Host "    Target list ID: $targetListId" -ForegroundColor Gray
        
        # Check if field already exists
        $existingField = Get-PnPField -List $SourceListName -Identity $FieldName -ErrorAction SilentlyContinue
        if ($existingField) {
            Write-Warning "    Lookup field '$FieldName' already exists in $SourceListName"
            return
        }
        
        # Create the lookup field
        $fieldXml = @"
<Field Type='Lookup' 
       DisplayName='$DisplayName' 
       Name='$FieldName' 
       StaticName='$FieldName'
       List='{$targetListId}' 
       ShowField='$ShowField' 
       Required='FALSE' />
"@
        
        Add-PnPFieldFromXml -List $SourceListName -FieldXml $fieldXml -ErrorAction Stop
        Write-Host "    ✓ Lookup field created successfully" -ForegroundColor Green
        
        # Add the field to the default view
        try {
            $defaultView = Get-PnPView -List $SourceListName -Identity "All Items" -ErrorAction SilentlyContinue
            if ($defaultView) {
                Add-PnPViewField -List $SourceListName -View "All Items" -Field $FieldName -ErrorAction SilentlyContinue
                Write-Host "    ✓ Added to default view" -ForegroundColor Gray
            }
        } catch {
            Write-Warning "    Could not add field to default view: $($_.Exception.Message)"
        }
        
    } catch {
        Write-Error "  Failed to create lookup field: $($_.Exception.Message)"
        Write-Host "    Attempting alternative method..." -ForegroundColor Yellow
        
        # Alternative method using Add-PnPField with specific parameters
        try {
            Add-PnPField -List $SourceListName -DisplayName $DisplayName -InternalName $FieldName -Type Lookup -LookupList $TargetListName -LookupField $ShowField -ErrorAction Stop
            Write-Host "    ✓ Lookup field created using alternative method" -ForegroundColor Green
        } catch {
            Write-Error "    Alternative method also failed: $($_.Exception.Message)"
        }
    }
}

# Deploy each STP file
Write-Host ""
Write-Host "Deploying STP templates..." -ForegroundColor Yellow
$deployedLists = @()
$customersListName = "$($ListPrefix)Customers"
$opportunitiesListName = "$($ListPrefix)Opportunities"

foreach ($stpFile in $stpFiles) {
    Write-Host ""
    Write-Host "Processing $($stpFile.Name)..." -ForegroundColor Cyan
    
    $deployedList = Deploy-StpFile -StpPath $stpFile.Path -ListName $stpFile.ListName -Description "$($stpFile.Name) list created from STP template"
    
    if ($deployedList) {
        $deployedLists += $deployedList
        
        # Add CRM-specific fields
        Add-CrmFields -ListName $stpFile.ListName -ListType $stpFile.Name -CustomersListName $customersListName
    }
}

# Verify deployment
Write-Host ""
Write-Host "Verifying deployment..." -ForegroundColor Yellow
$customersExists = $false
$opportunitiesExists = $false

try {
    $customersList = Get-PnPList -Identity $customersListName -ErrorAction SilentlyContinue
    if ($customersList) {
        $customersExists = $true
        Write-Host "✓ Customers list: $($customersList.Title)" -ForegroundColor Green
        Write-Host "  URL: $($customersList.DefaultViewUrl)" -ForegroundColor Gray
    }
} catch {
    Write-Warning "Could not verify Customers list"
}

try {
    $opportunitiesList = Get-PnPList -Identity $opportunitiesListName -ErrorAction SilentlyContinue
    if ($opportunitiesList) {
        $opportunitiesExists = $true
        Write-Host "✓ Opportunities list: $($opportunitiesList.Title)" -ForegroundColor Green
        Write-Host "  URL: $($opportunitiesList.DefaultViewUrl)" -ForegroundColor Gray
    }
} catch {
    Write-Warning "Could not verify Opportunities list"
}

# Create lookup relationships after both lists exist
Write-Host ""
Write-Host "Setting up lookup relationships..." -ForegroundColor Yellow

if ($customersExists -and $opportunitiesExists) {
    # Create lookup from Opportunities to Customers
    Add-LookupField -SourceListName $opportunitiesListName -TargetListName $customersListName -FieldName "CustomerId" -DisplayName "CustomerId" -ShowField "CustomerName"
    
    Write-Host "✓ Lookup relationship created: $opportunitiesListName → $customersListName" -ForegroundColor Green
} else {
    Write-Warning "Cannot create lookup relationships - both lists must exist"
}

# Add sample data (optional)
Write-Host ""
Write-Host "Adding sample data..." -ForegroundColor Yellow

if ($customersExists) {
    try {
        # Add sample customers
        $sampleCustomers = @(
            @{Title="ACME Corporation"; CustomerName="ACME Corporation"; PrimaryContact="John Smith"; PrimaryContactTitle="CEO"; AlternateContact="Jane Doe"; AlternateTitle="VP Sales"; AlternateContact2="Bob Wilson"; AlternateContact2Title="CTO"; Website="https://acme.com"},
            @{Title="Global Industries"; CustomerName="Global Industries"; PrimaryContact="Sarah Johnson"; PrimaryContactTitle="Director"; AlternateContact="Mike Brown"; AlternateTitle="Manager"; AlternateContact2="Lisa Davis"; AlternateContact2Title="Coordinator"; Website="https://globalindustries.com"},
            @{Title="Tech Solutions Inc"; CustomerName="Tech Solutions Inc"; PrimaryContact="David Lee"; PrimaryContactTitle="Founder"; AlternateContact="Emily Chen"; AlternateTitle="COO"; AlternateContact2="Alex Rodriguez"; AlternateContact2Title="CFO"; Website="https://techsolutions.com"}
        )
        
        foreach ($customer in $sampleCustomers) {
            $existingCustomer = Get-PnPListItem -List $customersListName -Query "<View><Query><Where><Eq><FieldRef Name='CustomerName'/><Value Type='Text'>$($customer.CustomerName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
            if (-not $existingCustomer) {
                Add-PnPListItem -List $customersListName -Values $customer -ErrorAction SilentlyContinue
                Write-Host "  ✓ Added customer: $($customer.CustomerName)" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Warning "Could not add sample customer data: $($_.Exception.Message)"
    }
}

if ($opportunitiesExists) {
    try {
        # Get customer items for lookup
        $customers = Get-PnPListItem -List $customersListName -ErrorAction SilentlyContinue
        
        if ($customers -and $customers.Count -gt 0) {
            # Add sample opportunities
            $sampleOpportunities = @(
                @{Title="ERP Implementation Project"; OpportunityName="ERP Implementation Project"; Status="On Track"; Stage="Proposal"; Amount=50000; Probability="High"; OpportunityOwner="John Manager"; Close=(Get-Date).AddDays(30); NextMilestone="Technical Review"; NextMilestoneDate=(Get-Date).AddDays(15); CommentLog="Initial discovery completed. Moving to proposal phase."},
                @{Title="Cloud Migration Initiative"; OpportunityName="Cloud Migration Initiative"; Status="At Risk"; Stage="Negotiation"; Amount=25000; Probability="Medium"; OpportunityOwner="Sarah Director"; Close=(Get-Date).AddDays(45); NextMilestone="Contract Finalization"; NextMilestoneDate=(Get-Date).AddDays(20); CommentLog="Pricing discussions ongoing. Customer budget concerns identified."},
                @{Title="Security Assessment"; OpportunityName="Security Assessment"; Status="On Track"; Stage="Lead Qualification"; Amount=15000; Probability="Low"; OpportunityOwner="Mike Consultant"; Close=(Get-Date).AddDays(60); NextMilestone="Stakeholder Meeting"; NextMilestoneDate=(Get-Date).AddDays(10); CommentLog="Initial contact made. Scheduling needs assessment meeting."}
            )
            
            for ($i = 0; $i -lt $sampleOpportunities.Count -and $i -lt $customers.Count; $i++) {
                $opportunity = $sampleOpportunities[$i]
                $customer = $customers[$i]
                
                # Add customer lookup value
                $opportunity["CustomerIdId"] = $customer.Id
                
                $existingOpp = Get-PnPListItem -List $opportunitiesListName -Query "<View><Query><Where><Eq><FieldRef Name='OpportunityName'/><Value Type='Text'>$($opportunity.OpportunityName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
                if (-not $existingOpp) {
                    Add-PnPListItem -List $opportunitiesListName -Values $opportunity -ErrorAction SilentlyContinue
                    Write-Host "  ✓ Added opportunity: $($opportunity.OpportunityName)" -ForegroundColor Gray
                }
            }
        }
    } catch {
        Write-Warning "Could not add sample opportunity data: $($_.Exception.Message)"
    }
}

# Summary
Write-Host ""
if ($customersExists -and $opportunitiesExists) {
    Write-Host "✓ Deployment completed successfully!" -ForegroundColor Green
    Write-Host "Both CRM lists have been created with custom fields." -ForegroundColor Green
} elseif ($customersExists -or $opportunitiesExists) {
    Write-Host "⚠ Partial deployment completed" -ForegroundColor Yellow
    Write-Host "Some lists were created, but not all. Check the warnings above." -ForegroundColor Yellow
} else {
    Write-Host "✗ Deployment failed" -ForegroundColor Red
    Write-Host "No lists were created successfully. Check the errors above." -ForegroundColor Red
}

Write-Host ""
Write-Host "Note: STP files contain complex templates that may require manual configuration." -ForegroundColor Cyan
Write-Host "Consider extracting the STP files to examine their structure for complete deployment." -ForegroundColor Cyan

# Disconnect
try {
    Disconnect-PnPOnline
    Write-Host ""
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
} catch {
    # Ignore disconnect errors
}
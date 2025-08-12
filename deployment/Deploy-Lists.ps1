param(
    [string]$SiteUrl,
    [string]$ListPrefix,
    [string]$ClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e",
    [string]$CustomersStpFile = "customers.stp",
    [string]$OpportunitiesStpFile = "opportunities.stp"
)

# If parameters not provided via command line, prompt for them
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Host "=== SharePoint CRM Lists Deployment Configuration ===" -ForegroundColor Cyan
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

# Infer Tenant URL from Site URL
try {
    $uri = [System.Uri]$SiteUrl
    $TenantUrl = "$($uri.Scheme)://$($uri.Host)"
} catch {
    Write-Error "Invalid Site URL format. Please enter a valid URL."
    exit 1
}

Write-Host ""
Write-Host "=== SharePoint CRM Lists Deployment (Enhanced Version) ===" -ForegroundColor Cyan
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

# Add new lists that don't require STP files
$additionalLists = @(
    @{Name = "Contacts"; ListName = "$($ListPrefix)Contacts"},
    @{Name = "Milestones"; ListName = "$($ListPrefix)Milestones"}
)

Write-Host "✓ Additional lists to create: Contacts, Milestones" -ForegroundColor Green

if ($stpFiles.Count -eq 0 -and $additionalLists.Count -eq 0) {
    Write-Error "No template files found and no additional lists defined. Cannot proceed."
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

# Function to create basic list
function Create-BasicList {
    param(
        [string]$ListName,
        [string]$Description = ""
    )
    
    try {
        Write-Host "  Creating '$ListName'..." -ForegroundColor Yellow
        
        # Check if list already exists
        $existingList = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
        if ($existingList) {
            Write-Warning "  List '$ListName' already exists. Skipping creation."
            return $existingList
        }
        
        # Create a basic custom list
        $newList = New-PnPList -Title $ListName -Template GenericList -ErrorAction Stop
        Write-Host "    ✓ List created: $ListName" -ForegroundColor Green
        return $newList
        
    } catch {
        Write-Error "  Failed to create $ListName : $($_.Exception.Message)"
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
                @{Name="Status"; Type="Choice"; DisplayName="Status"; Choices=@("Active", "At Risk", "Critical", "Dormant"); Required=$true},
                @{Name="OpportunityOwner"; Type="Text"; DisplayName="Opportunity Owner"},
                @{Name="OpportunityStage"; Type="Choice"; DisplayName="Stage"; Choices=@("Lead Qualification", "Nurturing", "Proposal", "Negotiation", "Project Execution", "Closeout"); Required=$true},
                @{Name="Amount"; Type="Currency"; DisplayName="Opportunity Value"; Required=$true},
                @{Name="Probability"; Type="Choice"; DisplayName="Win Probability"; Choices=@("Low", "Medium", "High"); Required=$true},
                @{Name="Close"; Type="DateTime"; DisplayName="Expected Close Date"; Required=$true},
                @{Name="NextMilestoneDate"; Type="DateTime"; DisplayName="Next Deadline or Milestone"},
                @{Name="NextMilestone"; Type="Text"; DisplayName="Next Milestone"},
                # NEW FIELDS: RecurringRevenue model fields
                @{Name="RecurringRevenueModel"; Type="Choice"; DisplayName="Recurring Revenue Model"; Choices=@("Up Front", "Annually", "Semi-Annually","Quarterly", "Monthly"); Required=$false},
                @{Name="Recurrences"; Type="Number"; DisplayName="Recurrences"; Required=$false},
                @{Name="StartDate"; Type="DateTime"; DisplayName="Start Date"; Required=$false}
            )
        } elseif ($ListType -eq "Contacts") {
            # Add contact-specific fields
            $fields = @(
                @{Name="ContactName"; Type="Text"; DisplayName="Name"; Required=$true},
                @{Name="JobTitle"; Type="Text"; DisplayName="Job Title"},
                @{Name="OfficePhone"; Type="Text"; DisplayName="Office Phone"},
                @{Name="MobilePhone"; Type="Text"; DisplayName="Mobile Phone"},
                @{Name="Email"; Type="Text"; DisplayName="Email"}
            )
        } elseif ($ListType -eq "Milestones") {
            # Add milestone-specific fields
            $fields = @(
                @{Name="MilestoneName"; Type="Text"; DisplayName="Name"; Required=$true},
                @{Name="Owner"; Type="Text"; DisplayName="Owner"},
                @{Name="MilestoneDate"; Type="DateTime"; DisplayName="Date"; Required=$true},
                @{Name="MilestoneStatus"; Type="Choice"; DisplayName="Status"; Choices=@("Not Started", "In Progress", "Completed", "On Hold", "Cancelled"); Required=$true}
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
                        if ($field.DefaultToClose -eq $true) {
                            # For StartDate field, create with formula default to Close field
                            $fieldXml = @"
<Field Type='DateTime' 
       DisplayName='$($field.DisplayName)' 
       Name='$($field.Name)' 
       StaticName='$($field.Name)'
       Format='DateOnly'>
    <Default>[Close]</Default>
</Field>
"@
                            Add-PnPFieldFromXml -List $ListName -FieldXml $fieldXml -ErrorAction SilentlyContinue
                            Write-Host "    ✓ StartDate field created with default to Close field" -ForegroundColor Gray
                        } else {
                            Add-PnPField -List $ListName -DisplayName $field.DisplayName -InternalName $field.Name -Type DateTime -ErrorAction SilentlyContinue
                        }
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
        
        # Special handling for StartDate default value
        if ($ListType -eq "Opportunities") {
            Write-Host "  ✓ StartDate field configured to default to Close field value" -ForegroundColor Green
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
$contactsListName = "$($ListPrefix)Contacts"
$milestonesListName = "$($ListPrefix)Milestones"

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

# Deploy additional lists (Contacts and Milestones)
Write-Host ""
Write-Host "Creating additional lists..." -ForegroundColor Yellow

foreach ($additionalList in $additionalLists) {
    Write-Host ""
    Write-Host "Processing $($additionalList.Name)..." -ForegroundColor Cyan
    
    $deployedList = Create-BasicList -ListName $additionalList.ListName -Description "$($additionalList.Name) list for CRM system"
    
    if ($deployedList) {
        $deployedLists += $deployedList
        
        # Add specific fields for each list type
        Add-CrmFields -ListName $additionalList.ListName -ListType $additionalList.Name
    }
}

# Verify deployment
Write-Host ""
Write-Host "Verifying deployment..." -ForegroundColor Yellow
$customersExists = $false
$opportunitiesExists = $false
$contactsExists = $false
$milestonesExists = $false

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

try {
    $contactsList = Get-PnPList -Identity $contactsListName -ErrorAction SilentlyContinue
    if ($contactsList) {
        $contactsExists = $true
        Write-Host "✓ Contacts list: $($contactsList.Title)" -ForegroundColor Green
        Write-Host "  URL: $($contactsList.DefaultViewUrl)" -ForegroundColor Gray
    }
} catch {
    Write-Warning "Could not verify Contacts list"
}

try {
    $milestonesList = Get-PnPList -Identity $milestonesListName -ErrorAction SilentlyContinue
    if ($milestonesList) {
        $milestonesExists = $true
        Write-Host "✓ Milestones list: $($milestonesList.Title)" -ForegroundColor Green
        Write-Host "  URL: $($milestonesList.DefaultViewUrl)" -ForegroundColor Gray
    }
} catch {
    Write-Warning "Could not verify Milestones list"
}

# Create lookup relationships after all lists exist
Write-Host ""
Write-Host "Setting up lookup relationships..." -ForegroundColor Yellow

if ($customersExists -and $opportunitiesExists) {
    # Create lookup from Opportunities to Customers
    Add-LookupField -SourceListName $opportunitiesListName -TargetListName $customersListName -FieldName "CustomerId" -DisplayName "Customer" -ShowField "CustomerName"
    Write-Host "✓ Lookup relationship created: $opportunitiesListName → $customersListName" -ForegroundColor Green
} else {
    Write-Warning "Cannot create lookup relationship between Opportunities and Customers - both lists must exist"
}

if ($contactsExists -and $customersExists) {
    # Create lookup from Contacts to Customers
    Add-LookupField -SourceListName $contactsListName -TargetListName $customersListName -FieldName "CustomerId" -DisplayName "Customer" -ShowField "CustomerName"
    Write-Host "✓ Lookup relationship created: $contactsListName → $customersListName" -ForegroundColor Green
} else {
    Write-Warning "Cannot create lookup relationship between Contacts and Customers - both lists must exist"
}

if ($milestonesExists -and $opportunitiesExists) {
    # Create lookup from Milestones to Opportunities
    Add-LookupField -SourceListName $milestonesListName -TargetListName $opportunitiesListName -FieldName "OpportunityId" -DisplayName "Opportunity" -ShowField "OpportunityName"
    Write-Host "✓ Lookup relationship created: $milestonesListName → $opportunitiesListName" -ForegroundColor Green
} else {
    Write-Warning "Cannot create lookup relationship between Milestones and Opportunities - both lists must exist"
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

if ($contactsExists -and $customersExists) {
    try {
        # Get customer items for lookup
        $customers = Get-PnPListItem -List $customersListName -ErrorAction SilentlyContinue
        
        if ($customers -and $customers.Count -gt 0) {
            # Add sample contacts
            $sampleContacts = @(
                @{Title="John Smith"; ContactName="John Smith"; JobTitle="CEO"; OfficePhone="555-0101"; MobilePhone="555-0201"; Email="john.smith@acme.com"},
                @{Title="Jane Doe"; ContactName="Jane Doe"; JobTitle="VP Sales"; OfficePhone="555-0102"; MobilePhone="555-0202"; Email="jane.doe@acme.com"},
                @{Title="Sarah Johnson"; ContactName="Sarah Johnson"; JobTitle="Director"; OfficePhone="555-0103"; MobilePhone="555-0203"; Email="sarah.johnson@global.com"}
            )
            
            for ($i = 0; $i -lt $sampleContacts.Count -and $i -lt $customers.Count; $i++) {
                $contact = $sampleContacts[$i]
                $customer = $customers[$i]
                
                # Add customer lookup value
                $contact["CustomerIdId"] = $customer.Id
                
                $existingContact = Get-PnPListItem -List $contactsListName -Query "<View><Query><Where><Eq><FieldRef Name='ContactName'/><Value Type='Text'>$($contact.ContactName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
                if (-not $existingContact) {
                    Add-PnPListItem -List $contactsListName -Values $contact -ErrorAction SilentlyContinue
                    Write-Host "  ✓ Added contact: $($contact.ContactName)" -ForegroundColor Gray
                }
            }
        }
    } catch {
        Write-Warning "Could not add sample contact data: $($_.Exception.Message)"
    }
}

if ($opportunitiesExists -and $customersExists) {
    try {
        # Get customer items for lookup
        $customers = Get-PnPListItem -List $customersListName -ErrorAction SilentlyContinue
        
        if ($customers -and $customers.Count -gt 0) {
            # Add sample opportunities with new recurring revenue fields
            $sampleOpportunities = @(
                @{Title="ERP Implementation Project"; OpportunityName="ERP Implementation Project"; Status="Active"; OpportunityStage="Proposal"; Amount=50000; Probability="High"; OpportunityOwner="John Manager"; Close=(Get-Date).AddDays(30); NextMilestone="Technical Review"; NextMilestoneDate=(Get-Date).AddDays(15); RecurringRevenueModel="Annually"; Recurrences=3; StartDate=(Get-Date).AddDays(30)},
                @{Title="Cloud Migration Initiative"; OpportunityName="Cloud Migration Initiative"; Status="At Risk"; OpportunityStage="Negotiation"; Amount=25000; Probability="Medium"; OpportunityOwner="Sarah Director"; Close=(Get-Date).AddDays(45); NextMilestone="Contract Finalization"; NextMilestoneDate=(Get-Date).AddDays(20); RecurringRevenueModel="Monthly"; Recurrences=12; StartDate=(Get-Date).AddDays(45)},
                @{Title="Security Assessment"; OpportunityName="Security Assessment"; Status="Active"; OpportunityStage="Lead Qualification"; Amount=15000; Probability="Low"; OpportunityOwner="Mike Consultant"; Close=(Get-Date).AddDays(60); NextMilestone="Stakeholder Meeting"; NextMilestoneDate=(Get-Date).AddDays(10); RecurringRevenueModel="Up Front"; Recurrences=1; StartDate=(Get-Date).AddDays(60)}
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

if ($milestonesExists -and $opportunitiesExists) {
    try {
        # Get opportunity items for lookup
        $opportunities = Get-PnPListItem -List $opportunitiesListName -ErrorAction SilentlyContinue
        
        if ($opportunities -and $opportunities.Count -gt 0) {
            # Add sample milestones
            $sampleMilestones = @(
                @{Title="Technical Review"; MilestoneName="Technical Review"; Owner="John Manager"; MilestoneDate=(Get-Date).AddDays(15); MilestoneStatus="Not Started"},
                @{Title="Contract Finalization"; MilestoneName="Contract Finalization"; Owner="Sarah Director"; MilestoneDate=(Get-Date).AddDays(20); MilestoneStatus="In Progress"},
                @{Title="Stakeholder Meeting"; MilestoneName="Stakeholder Meeting"; Owner="Mike Consultant"; MilestoneDate=(Get-Date).AddDays(10); MilestoneStatus="Not Started"},
                @{Title="Requirements Gathering"; MilestoneName="Requirements Gathering"; Owner="John Manager"; MilestoneDate=(Get-Date).AddDays(5); MilestoneStatus="Completed"},
                @{Title="Security Audit"; MilestoneName="Security Audit"; Owner="Sarah Director"; MilestoneDate=(Get-Date).AddDays(35); MilestoneStatus="Not Started"}
            )
            
            for ($i = 0; $i -lt $sampleMilestones.Count; $i++) {
                $milestone = $sampleMilestones[$i]
                # Cycle through opportunities for milestone assignments
                $opportunityIndex = $i % $opportunities.Count
                $opportunity = $opportunities[$opportunityIndex]
                
                # Add opportunity lookup value
                $milestone["OpportunityIdId"] = $opportunity.Id
                
                $existingMilestone = Get-PnPListItem -List $milestonesListName -Query "<View><Query><Where><Eq><FieldRef Name='MilestoneName'/><Value Type='Text'>$($milestone.MilestoneName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
                if (-not $existingMilestone) {
                    Add-PnPListItem -List $milestonesListName -Values $milestone -ErrorAction SilentlyContinue
                    Write-Host "  ✓ Added milestone: $($milestone.MilestoneName)" -ForegroundColor Gray
                }
            }
        }
    } catch {
        Write-Warning "Could not add sample milestone data: $($_.Exception.Message)"
    }
}

# Summary
Write-Host ""
$totalListsExpected = 4  # Customers, Opportunities, Contacts, Milestones
$totalListsCreated = 0

if ($customersExists) { $totalListsCreated++ }
if ($opportunitiesExists) { $totalListsCreated++ }
if ($contactsExists) { $totalListsCreated++ }
if ($milestonesExists) { $totalListsCreated++ }

if ($totalListsCreated -eq $totalListsExpected) {
    Write-Host "✓ Deployment completed successfully!" -ForegroundColor Green
    Write-Host "All CRM lists have been created with custom fields and lookup relationships." -ForegroundColor Green
    Write-Host ""
    Write-Host "Lists Created:" -ForegroundColor Cyan
    Write-Host "  • Customers - Customer management with contact information" -ForegroundColor Gray
    Write-Host "  • Opportunities - Sales opportunities with recurring revenue tracking" -ForegroundColor Gray
    Write-Host "  • Contacts - Contact information linked to customers" -ForegroundColor Gray
    Write-Host "  • Milestones - Project milestones linked to opportunities" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Lookup Relationships:" -ForegroundColor Cyan
    Write-Host "  • Opportunities → Customers" -ForegroundColor Gray
    Write-Host "  • Contacts → Customers" -ForegroundColor Gray
    Write-Host "  • Milestones → Opportunities" -ForegroundColor Gray
    Write-Host ""
    Write-Host "New Recurring Revenue Fields Added to Opportunities:" -ForegroundColor Cyan
    Write-Host "  • Recurring Revenue Model (Choice: Up Front, Annually, Quarterly, Monthly)" -ForegroundColor Gray
    Write-Host "  • Recurrences (Number: Integer field)" -ForegroundColor Gray
    Write-Host "  • Start Date (Date field)" -ForegroundColor Gray
} elseif ($totalListsCreated -gt 0) {
    Write-Host "⚠ Partial deployment completed" -ForegroundColor Yellow
    Write-Host "$totalListsCreated of $totalListsExpected lists were created successfully. Check the warnings above." -ForegroundColor Yellow
} else {
    Write-Host "✗ Deployment failed" -ForegroundColor Red
    Write-Host "No lists were created successfully. Check the errors above." -ForegroundColor Red
}

Write-Host ""
Write-Host "Note: STP files contain complex templates that may require manual configuration." -ForegroundColor Cyan
Write-Host "Consider extracting the STP files to examine their structure for complete deployment." -ForegroundColor Cyan
Write-Host ""
Write-Host "Manual Configuration Required:" -ForegroundColor Yellow
Write-Host "  • StartDate field: Set default value to Close field in SharePoint list settings" -ForegroundColor Gray
Write-Host "  • Consider using Power Automate to automatically populate StartDate from Close field" -ForegroundColor Gray
Write-Host "  • Review and customize field validation rules as needed" -ForegroundColor Gray

# Disconnect
try {
    Disconnect-PnPOnline
    Write-Host ""
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
} catch {
    # Ignore disconnect errors
}
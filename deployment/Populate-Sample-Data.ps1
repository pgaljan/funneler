param(
    [string]$SiteUrl,
    [string]$ListPrefix,
    [string]$ClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e"
)

# If parameters not provided via command line, prompt for them
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Host "=== SharePoint Sample Data Population ===" -ForegroundColor Cyan
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
Write-Host "=== SharePoint Sample Data Population ===" -ForegroundColor Cyan
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

# Function to add sample customers
function Add-SampleCustomers {
    param([string]$ListName)
    
    Write-Host "Adding sample customers to $ListName..." -ForegroundColor Yellow
    
    $sampleCustomers = @(
        @{
            Title = "ACME Corporation"
            CustomerName = "ACME Corporation"
            PrimaryContact = "John Smith"
            PrimaryContactTitle = "CEO"
            AlternateContact = "Jane Doe"
            AlternateContactTitle = "VP Sales"
            AlternateContact2 = "Bob Wilson"
            AlternateContact2Title = "CTO"
            Website = "https://acme.com"
            NAICSSector = "541511"
            CustomerStatus = "Active"
        },
        @{
            Title = "Global Industries"
            CustomerName = "Global Industries"
            PrimaryContact = "Sarah Johnson"
            PrimaryContactTitle = "Director"
            AlternateContact = "Mike Brown"
            AlternateContactTitle = "Manager"
            AlternateContact2 = "Lisa Davis"
            AlternateContact2Title = "Coordinator"
            Website = "https://globalindustries.com"
            NAICSSector = "336411"
            CustomerStatus = "Prospect"
        },
        @{
            Title = "Tech Solutions Inc"
            CustomerName = "Tech Solutions Inc"
            PrimaryContact = "David Lee"
            PrimaryContactTitle = "Founder"
            AlternateContact = "Emily Chen"
            AlternateContactTitle = "COO"
            AlternateContact2 = "Alex Rodriguez"
            AlternateContact2Title = "CFO"
            Website = "https://techsolutions.com"
            NAICSSector = "541512"
            CustomerStatus = "Active"
        },
        @{
            Title = "Manufacturing Plus"
            CustomerName = "Manufacturing Plus"
            PrimaryContact = "Robert Taylor"
            PrimaryContactTitle = "Plant Manager"
            AlternateContact = "Jennifer White"
            AlternateContactTitle = "Operations Director"
            AlternateContact2 = "Mark Thompson"
            AlternateContact2Title = "Quality Manager"
            Website = "https://manufacturingplus.com"
            NAICSSector = "332710"
            CustomerStatus = "Active"
        },
        @{
            Title = "Healthcare Partners"
            CustomerName = "Healthcare Partners"
            PrimaryContact = "Dr. Maria Garcia"
            PrimaryContactTitle = "Chief Medical Officer"
            AlternateContact = "James Anderson"
            AlternateContactTitle = "Administrator"
            AlternateContact2 = "Susan Miller"
            AlternateContact2Title = "Finance Director"
            Website = "https://healthcarepartners.com"
            NAICSSector = "621111"
            CustomerStatus = "Prospect"
        }
    )
    
    $addedCount = 0
    foreach ($customer in $sampleCustomers) {
        try {
            # Check if customer already exists
            $existingCustomer = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='CustomerName'/><Value Type='Text'>$($customer.CustomerName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
            
            if (-not $existingCustomer) {
                Add-PnPListItem -List $ListName -Values $customer -ErrorAction Stop
                Write-Host "  ✓ Added customer: $($customer.CustomerName)" -ForegroundColor Green
                $addedCount++
            } else {
                Write-Host "  - Customer already exists: $($customer.CustomerName)" -ForegroundColor Gray
            }
        } catch {
            Write-Warning "  Failed to add customer $($customer.CustomerName): $($_.Exception.Message)"
        }
    }
    
    Write-Host "  ✓ Added $addedCount new customers" -ForegroundColor Green
    return $addedCount
}

# Function to add sample opportunities
function Add-SampleOpportunities {
    param(
        [string]$ListName,
        [string]$CustomersListName
    )
    
    Write-Host "Adding sample opportunities to $ListName..." -ForegroundColor Yellow
    
    # Wait a moment for any recent field creation to propagate
    Write-Host "  Waiting for field propagation..." -ForegroundColor Gray
    Start-Sleep -Seconds 3
    
    # Check what the actual CustomerID lookup field internal name is
    $customerFieldName = $null
    $possibleFieldNames = @("CustomerID", "CustomerId", "CustomerIDId", "CustomerIdId", "CustomerIDLookupId")
    
    foreach ($fieldName in $possibleFieldNames) {
        try {
            $customerField = Get-PnPField -List $ListName -Identity $fieldName -ErrorAction SilentlyContinue
            if ($customerField) {
                $customerFieldName = $customerField.InternalName
                Write-Host "  Found customer lookup field: $fieldName (Internal: $customerFieldName)" -ForegroundColor Gray
                break
            }
        } catch {
            # Continue to next field name
        }
    }
    
    if (-not $customerFieldName) {
        Write-Warning "CustomerID lookup field not found in $ListName. Trying to list all fields..."
        try {
            $allFields = Get-PnPField -List $ListName | Where-Object { $_.TypeDisplayName -eq "Lookup" -and $_.Title -like "*Customer*" }
            if ($allFields) {
                Write-Host "  Found customer-related lookup fields:" -ForegroundColor Gray
                foreach ($field in $allFields) {
                    Write-Host "    - $($field.Title) (Internal: $($field.InternalName))" -ForegroundColor Gray
                }
                # Use the first customer-related lookup field found
                $customerFieldName = $allFields[0].InternalName
                Write-Host "  Using field: $customerFieldName" -ForegroundColor Yellow
            } else {
                Write-Warning "No customer lookup fields found. Please ensure the CustomerID lookup field exists in the opportunities list."
                return 0
            }
        } catch {
            Write-Error "Could not analyze fields in $ListName : $($_.Exception.Message)"
            return 0
        }
    }
    
    # Get customer items for lookup
    try {
        $customers = Get-PnPListItem -List $CustomersListName -ErrorAction Stop
        if ($customers.Count -eq 0) {
            Write-Warning "No customers found in $CustomersListName. Cannot create opportunities."
            return 0
        }
    } catch {
        Write-Error "Failed to get customers from $CustomersListName : $($_.Exception.Message)"
        return 0
    }
    
    $sampleOpportunities = @(
        @{
            Title = "ERP Implementation Project"
            OpportunityName = "ERP Implementation Project"
            Status = "Active"
            OpportunityStage = "Proposal"
            Amount = 125000
            Probability = "High"
            OpportunityOwner = "John Manager"
            Close = (Get-Date).AddDays(45)
            NextMilestone = "Technical Review"
            NextMilestoneDate = (Get-Date).AddDays(15)
        },
        @{
            Title = "Cloud Migration Initiative"
            OpportunityName = "Cloud Migration Initiative"
            Status = "At Risk"
            OpportunityStage = "Negotiation"
            Amount = 75000
            Probability = "Medium"
            OpportunityOwner = "Sarah Director"
            Close = (Get-Date).AddDays(60)
            NextMilestone = "Contract Finalization"
            NextMilestoneDate = (Get-Date).AddDays(20)
        },
        @{
            Title = "Security Assessment & Audit"
            OpportunityName = "Security Assessment & Audit"
            Status = "Active"
            OpportunityStage = "Lead Qualification"
            Amount = 35000
            Probability = "Low"
            OpportunityOwner = "Mike Consultant"
            Close = (Get-Date).AddDays(90)
            NextMilestone = "Stakeholder Meeting"
            NextMilestoneDate = (Get-Date).AddDays(10)
        },
        @{
            Title = "Manufacturing Automation System"
            OpportunityName = "Manufacturing Automation System"
            Status = "Active"
            OpportunityStage = "Proposal"
            Amount = 250000
            Probability = "High"
            OpportunityOwner = "Lisa Engineer"
            Close = (Get-Date).AddDays(120)
            NextMilestone = "Site Survey"
            NextMilestoneDate = (Get-Date).AddDays(30)
        },
        @{
            Title = "Healthcare Data Analytics Platform"
            OpportunityName = "Healthcare Data Analytics Platform"
            Status = "Critical"
            OpportunityStage = "Nurturing"
            Amount = 180000
            Probability = "Medium"
            OpportunityOwner = "Dr. Tech Advisor"
            Close = (Get-Date).AddDays(75)
            NextMilestone = "ROI Presentation"
            NextMilestoneDate = (Get-Date).AddDays(14)
        }
    )
    
    $addedCount = 0
    for ($i = 0; $i -lt $sampleOpportunities.Count -and $i -lt $customers.Count; $i++) {
        try {
            $opportunity = $sampleOpportunities[$i]
            $customer = $customers[$i]
            
            # Add customer lookup value using the correct field name
            $opportunity[$customerFieldName] = $customer.Id
            
            # Check if opportunity already exists
            $existingOpp = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='OpportunityName'/><Value Type='Text'>$($opportunity.OpportunityName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
            
            if (-not $existingOpp) {
                Add-PnPListItem -List $ListName -Values $opportunity -ErrorAction Stop
                Write-Host "  ✓ Added opportunity: $($opportunity.OpportunityName)" -ForegroundColor Green
                $addedCount++
            } else {
                Write-Host "  - Opportunity already exists: $($opportunity.OpportunityName)" -ForegroundColor Gray
            }
        } catch {
            Write-Warning "  Failed to add opportunity $($opportunity.OpportunityName): $($_.Exception.Message)"
        }
    }
    
    Write-Host "  ✓ Added $addedCount new opportunities" -ForegroundColor Green
    return $addedCount
}

# Check if lists exist
Write-Host ""
Write-Host "Checking for existing lists..." -ForegroundColor Yellow

$customersExists = Test-ListExists -ListName $customersListName
$opportunitiesExists = Test-ListExists -ListName $opportunitiesListName

if ($customersExists) {
    Write-Host "✓ Found Customers list: $customersListName" -ForegroundColor Green
} else {
    Write-Warning "Customers list '$customersListName' not found"
}

if ($opportunitiesExists) {
    Write-Host "✓ Found Opportunities list: $opportunitiesListName" -ForegroundColor Green
} else {
    Write-Warning "Opportunities list '$opportunitiesListName' not found"
}

if (-not $customersExists -and -not $opportunitiesExists) {
    Write-Error "Neither list found. Please check the list prefix and ensure lists have been created."
    Disconnect-PnPOnline
    exit 1
}

# Populate sample data
Write-Host ""
Write-Host "Populating sample data..." -ForegroundColor Yellow

$totalCustomersAdded = 0
$totalOpportunitiesAdded = 0

# Add customers
if ($customersExists) {
    Write-Host ""
    $totalCustomersAdded = Add-SampleCustomers -ListName $customersListName
}

# Add opportunities (requires customers to exist for lookups)
if ($opportunitiesExists -and $customersExists) {
    Write-Host ""
    $totalOpportunitiesAdded = Add-SampleOpportunities -ListName $opportunitiesListName -CustomersListName $customersListName
} elseif ($opportunitiesExists -and -not $customersExists) {
    Write-Warning "Cannot add opportunities without customers list for lookup relationships"
}

# Summary
Write-Host ""
Write-Host "=== Sample Data Population Complete ===" -ForegroundColor Green

if ($customersExists) {
    Write-Host "✓ Customers list: $totalCustomersAdded new records added" -ForegroundColor Gray
    Write-Host "  Companies: ACME Corp, Global Industries, Tech Solutions Inc," -ForegroundColor Gray
    Write-Host "             Manufacturing Plus, Healthcare Partners" -ForegroundColor Gray
}

if ($opportunitiesExists) {
    Write-Host "✓ Opportunities list: $totalOpportunitiesAdded new records added" -ForegroundColor Gray
    Write-Host "  Projects: ERP Implementation, Cloud Migration, Security Audit," -ForegroundColor Gray
    Write-Host "            Manufacturing Automation, Healthcare Analytics" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Sample data includes:" -ForegroundColor Cyan
Write-Host "  • Various customer statuses (Active, Prospect)" -ForegroundColor Gray
Write-Host "  • Different opportunity stages and values" -ForegroundColor Gray
Write-Host "  • Realistic contact information and dates" -ForegroundColor Gray
Write-Host "  • Proper lookup relationships between lists" -ForegroundColor Gray

# Disconnect
try {
    Disconnect-PnPOnline
    Write-Host ""
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
} catch {
    # Ignore disconnect errors
}
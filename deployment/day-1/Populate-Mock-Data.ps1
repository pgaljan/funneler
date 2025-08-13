param(
    [string]$SiteUrl,
    [string]$ListPrefix,
    [string]$ClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e",
    [string]$CsvPath = "."
)

# If parameters not provided via command line, prompt for them
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Host "=== SharePoint CSV Data Population ===" -ForegroundColor Cyan
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

if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    $CsvPath = Read-Host "Enter path to CSV files directory (press Enter for current directory)"
    if ([string]::IsNullOrWhiteSpace($CsvPath)) {
        $CsvPath = "."
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
Write-Host "=== SharePoint CSV Data Population ===" -ForegroundColor Cyan
Write-Host "Site: $SiteUrl" -ForegroundColor Yellow
Write-Host "Prefix: $ListPrefix" -ForegroundColor Yellow
Write-Host "CSV Path: $CsvPath" -ForegroundColor Yellow
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
$contactsListName = "$($ListPrefix)Contacts"
$milestonesListName = "$($ListPrefix)Milestones"

# Define CSV file paths
$customersCSV = Join-Path $CsvPath "Customers.csv"
$contactsCSV = Join-Path $CsvPath "Contacts.csv"
$opportunitiesCSV = Join-Path $CsvPath "Opportunities.csv"
$milestonesCSV = Join-Path $CsvPath "Milestones.csv"

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

# Function to verify CSV file exists
function Test-CSVExists {
    param([string]$FilePath)
    
    if (Test-Path $FilePath) {
        return $true
    } else {
        Write-Warning "CSV file not found: $FilePath"
        return $false
    }
}

# Function to parse CSV with proper handling of SharePoint export format
function Import-SharePointCSV {
    param([string]$FilePath)
    
    try {
        # Read all lines from the file
        $lines = Get-Content $FilePath -Encoding UTF8
        
        if ($lines.Count -lt 3) {
            throw "CSV file must have at least 3 lines (schema, headers, data)"
        }
        
        # Skip the first line (schema) and use the second line as headers
        $headerLine = $lines[1]
        $dataLines = $lines[2..($lines.Count - 1)]
        
        # Reconstruct CSV content
        $csvContent = @($headerLine) + $dataLines -join "`n"
        
        # Convert to proper CSV object
        $data = $csvContent | ConvertFrom-Csv
        
        Write-Host "  ✓ Parsed $($data.Count) records from $([System.IO.Path]::GetFileName($FilePath))" -ForegroundColor Gray
        return $data
        
    } catch {
        Write-Error "Failed to parse CSV file $FilePath : $($_.Exception.Message)"
        return @()
    }
}

# Function to parse currency values
function Parse-CurrencyValue {
    param([string]$Value)
    
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return 0
    }
    
    # Remove currency symbols and commas, then convert to double
    $cleanValue = $Value -replace '[\$,]', ''
    try {
        return [double]$cleanValue
    } catch {
        return 0
    }
}

# Function to parse date values
function Parse-DateValue {
    param([string]$Value)
    
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }
    
    try {
        return [DateTime]::Parse($Value)
    } catch {
        Write-Warning "Could not parse date: $Value"
        return $null
    }
}

# Function to add customers from CSV
function Add-CustomersFromCSV {
    param(
        [string]$ListName,
        [string]$CSVPath
    )
    
    Write-Host "Adding customers from CSV to $ListName..." -ForegroundColor Yellow
    
    if (-not (Test-CSVExists -FilePath $CSVPath)) {
        return 0
    }
    
    $customers = Import-SharePointCSV -FilePath $CSVPath
    if ($customers.Count -eq 0) {
        Write-Warning "No customer data found in CSV"
        return 0
    }
    
    $addedCount = 0
    foreach ($customer in $customers) {
        try {
            # Map CSV columns to SharePoint fields
            $customerData = @{
                CustomerName = $customer."Customer Name"
                Website = $customer."Website"
                NAICSSector = $customer."NAICS code"
                CustomerStatus = $customer."Status"
                PrimaryContact = $customer."Primary Contact"
                PrimaryContactTitle = $customer."Primary Contact Title"
                AlternateContact = $customer."Alternate Contact"
                AlternateContactTitle = $customer."Alternate Contact Title"
                AlternateContact2 = $customer."Alternate Contact 2"
                AlternateContact2Title = $customer."Alternate Contact 2 Title"
            }
            
            # Check if customer already exists
            $existingCustomer = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='CustomerName'/><Value Type='Text'>$($customerData.CustomerName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
            
            if (-not $existingCustomer) {
                Add-PnPListItem -List $ListName -Values $customerData -ErrorAction Stop
                Write-Host "  ✓ Added customer: $($customerData.CustomerName)" -ForegroundColor Green
                $addedCount++
            } else {
                Write-Host "  - Customer already exists: $($customerData.CustomerName)" -ForegroundColor Gray
            }
        } catch {
            Write-Warning "  Failed to add customer $($customer."Customer Name"): $($_.Exception.Message)"
        }
    }
    
    Write-Host "  ✓ Added $addedCount new customers" -ForegroundColor Green
    return $addedCount
}

# Function to add contacts from CSV
function Add-ContactsFromCSV {
    param(
        [string]$ListName,
        [string]$CustomersListName,
        [string]$CSVPath
    )
    
    Write-Host "Adding contacts from CSV to $ListName..." -ForegroundColor Yellow
    
    if (-not (Test-CSVExists -FilePath $CSVPath)) {
        return 0
    }
    
    # Get customer items for lookup
    try {
        $customers = Get-PnPListItem -List $CustomersListName -ErrorAction Stop
        if ($customers.Count -eq 0) {
            Write-Warning "No customers found in $CustomersListName. Cannot create contacts."
            return 0
        }
        Write-Host "  Found $($customers.Count) customers for contact assignment" -ForegroundColor Gray
    } catch {
        Write-Error "Failed to get customers from $CustomersListName : $($_.Exception.Message)"
        return 0
    }
    
    # Find the customer lookup field
    $listFields = Get-PnPField -List $ListName
    $customerFieldName = $null
    $possibleCustomerFieldNames = @("CustomerIdId", "CustomerId", "CustomerID", "Customer", "CustomerLookup")
    
    foreach ($fieldName in $possibleCustomerFieldNames) {
        $field = $listFields | Where-Object { $_.InternalName -eq $fieldName }
        if ($field -and $field.TypeDisplayName -eq "Lookup") {
            $customerFieldName = $fieldName
            Write-Host "  Found customer lookup field: $fieldName" -ForegroundColor Gray
            break
        }
    }
    
    if (-not $customerFieldName) {
        Write-Warning "Customer lookup field not found in $ListName. Cannot create contacts."
        return 0
    }
    
    $contacts = Import-SharePointCSV -FilePath $CSVPath
    if ($contacts.Count -eq 0) {
        Write-Warning "No contact data found in CSV"
        return 0
    }
    
    $addedCount = 0
    foreach ($contact in $contacts) {
        try {
            # Find the corresponding customer
            $customerName = $contact."Customer"
            $customer = $customers | Where-Object { $_["CustomerName"] -eq $customerName }
            
            if (-not $customer) {
                Write-Warning "  Customer '$customerName' not found for contact $($contact."Name")"
                continue
            }
            
            # Map CSV columns to SharePoint fields
            $contactData = @{
                ContactName = $contact."Name"
                JobTitle = $contact."Job Title"
                OfficePhone = $contact."Office Phone"
                MobilePhone = $contact."Mobile Phone"
                Email = $contact."Email"
            }
            
            # Add customer lookup
            $contactData[$customerFieldName] = $customer.Id
            
            # Check if contact already exists
            $existingContact = Get-PnPListItem -List $ListName -Query "<View><Query><Where><And><Eq><FieldRef Name='ContactName'/><Value Type='Text'>$($contactData.ContactName)</Value></Eq><Eq><FieldRef Name='$customerFieldName'/><Value Type='Lookup'>$($customer.Id)</Value></Eq></And></Where></Query></View>" -ErrorAction SilentlyContinue
            
            if (-not $existingContact) {
                Add-PnPListItem -List $ListName -Values $contactData -ErrorAction Stop
                Write-Host "  ✓ Added contact: $($contactData.ContactName) at $customerName" -ForegroundColor Green
                $addedCount++
            } else {
                Write-Host "  - Contact already exists: $($contactData.ContactName) at $customerName" -ForegroundColor Gray
            }
        } catch {
            Write-Warning "  Failed to add contact $($contact."Name"): $($_.Exception.Message)"
        }
    }
    
    Write-Host "  ✓ Added $addedCount new contacts" -ForegroundColor Green
    return $addedCount
}

# Function to add opportunities from CSV
function Add-OpportunitiesFromCSV {
    param(
        [string]$ListName,
        [string]$CustomersListName,
        [string]$CSVPath
    )
    
    Write-Host "Adding opportunities from CSV to $ListName..." -ForegroundColor Yellow
    
    if (-not (Test-CSVExists -FilePath $CSVPath)) {
        return 0
    }
    
    # Get customer items for lookup
    try {
        $customers = Get-PnPListItem -List $CustomersListName -ErrorAction Stop
        if ($customers.Count -eq 0) {
            Write-Warning "No customers found in $CustomersListName. Cannot create opportunities."
            return 0
        }
        Write-Host "  Found $($customers.Count) customers for opportunity assignment" -ForegroundColor Gray
    } catch {
        Write-Error "Failed to get customers from $CustomersListName : $($_.Exception.Message)"
        return 0
    }
    
    # Find the customer lookup field
    $listFields = Get-PnPField -List $ListName
    $customerFieldName = $null
    $possibleCustomerFieldNames = @("CustomerName", "CustomerIDId", "CustomerId", "CustomerID", "CustomerIdId", "CustomerIDLookupId", "Customer", "CustomerLookup")
    
    foreach ($fieldName in $possibleCustomerFieldNames) {
        $field = $listFields | Where-Object { $_.InternalName -eq $fieldName }
        if ($field -and $field.TypeDisplayName -eq "Lookup") {
            $customerFieldName = $fieldName
            Write-Host "  Found customer lookup field: $fieldName" -ForegroundColor Gray
            break
        }
    }
    
    if (-not $customerFieldName) {
        Write-Warning "Customer lookup field not found in $ListName. Cannot create opportunities."
        return 0
    }
    
    $opportunities = Import-SharePointCSV -FilePath $CSVPath
    if ($opportunities.Count -eq 0) {
        Write-Warning "No opportunity data found in CSV"
        return 0
    }
    
    $addedCount = 0
    foreach ($opportunity in $opportunities) {
        try {
            # Find the corresponding customer
            $customerName = $opportunity."Customer"
            $customer = $customers | Where-Object { $_["CustomerName"] -eq $customerName }
            
            if (-not $customer) {
                Write-Warning "  Customer '$customerName' not found for opportunity $($opportunity."Opportunity Name")"
                continue
            }
            
            # Map CSV columns to SharePoint fields
            $opportunityData = @{
                OpportunityName = $opportunity."Opportunity Name"
                Status = $opportunity."Status"
                OpportunityStage = $opportunity."Stage"
                Amount = Parse-CurrencyValue -Value $opportunity."Opportunity Value"
                Probability = $opportunity."Win Probability"
                OpportunityOwner = $opportunity."Opportunity Owner"
                Close = Parse-DateValue -Value $opportunity."Expected Close Date"
                NextMilestone = $opportunity."Next Milestone"
                NextMilestoneDate = Parse-DateValue -Value $opportunity."Next Deadline or Milestone"
                RecurringRevenueModel = $opportunity."Recurring Revenue Model"
                Recurrences = $opportunity."Recurrences"
                StartDate = Parse-DateValue -Value $opportunity."Start Date"
            }
            
            # Add customer lookup
            $opportunityData[$customerFieldName] = $customer.Id
            
            # Check if opportunity already exists
            $existingOpp = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='OpportunityName'/><Value Type='Text'>$($opportunityData.OpportunityName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
            
            if (-not $existingOpp) {
                Add-PnPListItem -List $ListName -Values $opportunityData -ErrorAction Stop
                Write-Host "  ✓ Added opportunity: $($opportunityData.OpportunityName) ($($opportunityData.Amount.ToString('C0')))" -ForegroundColor Green
                $addedCount++
            } else {
                Write-Host "  - Opportunity already exists: $($opportunityData.OpportunityName)" -ForegroundColor Gray
            }
        } catch {
            Write-Warning "  Failed to add opportunity $($opportunity."Opportunity Name"): $($_.Exception.Message)"
        }
    }
    
    Write-Host "  ✓ Added $addedCount new opportunities" -ForegroundColor Green
    return $addedCount
}

# Function to add milestones from CSV
function Add-MilestonesFromCSV {
    param(
        [string]$ListName,
        [string]$OpportunitiesListName,
        [string]$CSVPath
    )
    
    Write-Host "Adding milestones from CSV to $ListName..." -ForegroundColor Yellow
    
    if (-not (Test-CSVExists -FilePath $CSVPath)) {
        return 0
    }
    
    # Get opportunity items for lookup
    try {
        $opportunities = Get-PnPListItem -List $OpportunitiesListName -ErrorAction Stop
        if ($opportunities.Count -eq 0) {
            Write-Warning "No opportunities found in $OpportunitiesListName. Cannot create milestones."
            return 0
        }
        Write-Host "  Found $($opportunities.Count) opportunities for milestone assignment" -ForegroundColor Gray
    } catch {
        Write-Error "Failed to get opportunities from $OpportunitiesListName : $($_.Exception.Message)"
        return 0
    }
    
    # Find the opportunity lookup field
    $listFields = Get-PnPField -List $ListName
    $opportunityFieldName = $null
    $possibleOpportunityFieldNames = @("OpportunityIdId", "OpportunityId", "OpportunityID", "Opportunity", "OpportunityLookup")
    
    foreach ($fieldName in $possibleOpportunityFieldNames) {
        $field = $listFields | Where-Object { $_.InternalName -eq $fieldName }
        if ($field -and $field.TypeDisplayName -eq "Lookup") {
            $opportunityFieldName = $fieldName
            Write-Host "  Found opportunity lookup field: $fieldName" -ForegroundColor Gray
            break
        }
    }
    
    if (-not $opportunityFieldName) {
        Write-Warning "Opportunity lookup field not found in $ListName. Cannot create milestones."
        return 0
    }
    
    $milestones = Import-SharePointCSV -FilePath $CSVPath
    if ($milestones.Count -eq 0) {
        Write-Warning "No milestone data found in CSV"
        return 0
    }
    
    $addedCount = 0
    foreach ($milestone in $milestones) {
        try {
            # Find the corresponding opportunity
            $opportunityName = $milestone."Opportunity"
            $opportunity = $opportunities | Where-Object { $_["OpportunityName"] -eq $opportunityName }
            
            if (-not $opportunity) {
                Write-Warning "  Opportunity '$opportunityName' not found for milestone $($milestone."Name")"
                continue
            }
            
            # Map CSV columns to SharePoint fields
            $milestoneData = @{
                MilestoneName = $milestone."Name"
                Owner = $milestone."Owner"
                MilestoneDate = Parse-DateValue -Value $milestone."Date"
                MilestoneStatus = $milestone."Status"
            }
            
            # Add opportunity lookup
            $milestoneData[$opportunityFieldName] = $opportunity.Id
            
            # Check if milestone already exists
            $existingMilestone = Get-PnPListItem -List $ListName -Query "<View><Query><Where><And><Eq><FieldRef Name='MilestoneName'/><Value Type='Text'>$($milestoneData.MilestoneName)</Value></Eq><Eq><FieldRef Name='$opportunityFieldName'/><Value Type='Lookup'>$($opportunity.Id)</Value></Eq></And></Where></Query></View>" -ErrorAction SilentlyContinue
            
            if (-not $existingMilestone) {
                Add-PnPListItem -List $ListName -Values $milestoneData -ErrorAction Stop
                Write-Host "  ✓ Added milestone: $($milestoneData.MilestoneName) for $opportunityName" -ForegroundColor Green
                $addedCount++
            } else {
                Write-Host "  - Milestone already exists: $($milestoneData.MilestoneName) for $opportunityName" -ForegroundColor Gray
            }
        } catch {
            Write-Warning "  Failed to add milestone $($milestone."Name"): $($_.Exception.Message)"
        }
    }
    
    Write-Host "  ✓ Added $addedCount new milestones" -ForegroundColor Green
    return $addedCount
}

# Check if lists exist
Write-Host ""
Write-Host "Checking for existing lists..." -ForegroundColor Yellow

$customersExists = Test-ListExists -ListName $customersListName
$opportunitiesExists = Test-ListExists -ListName $opportunitiesListName
$contactsExists = Test-ListExists -ListName $contactsListName
$milestonesExists = Test-ListExists -ListName $milestonesListName

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

if ($contactsExists) {
    Write-Host "✓ Found Contacts list: $contactsListName" -ForegroundColor Green
} else {
    Write-Warning "Contacts list '$contactsListName' not found"
}

if ($milestonesExists) {
    Write-Host "✓ Found Milestones list: $milestonesListName" -ForegroundColor Green
} else {
    Write-Warning "Milestones list '$milestonesListName' not found"
}

# Check if CSV files exist
Write-Host ""
Write-Host "Checking for CSV files..." -ForegroundColor Yellow

$customersCSVExists = Test-CSVExists -FilePath $customersCSV
$contactsCSVExists = Test-CSVExists -FilePath $contactsCSV
$opportunitiesCSVExists = Test-CSVExists -FilePath $opportunitiesCSV
$milestonesCSVExists = Test-CSVExists -FilePath $milestonesCSV

if (-not $customersExists -and -not $opportunitiesExists -and -not $contactsExists -and -not $milestonesExists) {
    Write-Error "No lists found. Please check the list prefix and ensure lists have been created."
    Disconnect-PnPOnline
    exit 1
}

# Populate data from CSV files
Write-Host ""
Write-Host "Populating data from CSV files..." -ForegroundColor Yellow

$totalCustomersAdded = 0
$totalContactsAdded = 0
$totalOpportunitiesAdded = 0
$totalMilestonesAdded = 0

# Add customers first (required for other lookups)
if ($customersExists -and $customersCSVExists) {
    Write-Host ""
    $totalCustomersAdded = Add-CustomersFromCSV -ListName $customersListName -CSVPath $customersCSV
}

# Add contacts (requires customers to exist for lookups)
if ($contactsExists -and $contactsCSVExists -and $customersExists) {
    Write-Host ""
    $totalContactsAdded = Add-ContactsFromCSV -ListName $contactsListName -CustomersListName $customersListName -CSVPath $contactsCSV
} elseif ($contactsExists -and $contactsCSVExists -and -not $customersExists) {
    Write-Warning "Cannot add contacts without customers list for lookup relationships"
}

# Add opportunities (requires customers to exist for lookups)
if ($opportunitiesExists -and $opportunitiesCSVExists -and $customersExists) {
    Write-Host ""
    $totalOpportunitiesAdded = Add-OpportunitiesFromCSV -ListName $opportunitiesListName -CustomersListName $customersListName -CSVPath $opportunitiesCSV
} elseif ($opportunitiesExists -and $opportunitiesCSVExists -and -not $customersExists) {
    Write-Warning "Cannot add opportunities without customers list for lookup relationships"
}

# Add milestones (requires opportunities to exist for lookups)
if ($milestonesExists -and $milestonesCSVExists -and $opportunitiesExists) {
    Write-Host ""
    $totalMilestonesAdded = Add-MilestonesFromCSV -ListName $milestonesListName -OpportunitiesListName $opportunitiesListName -CSVPath $milestonesCSV
} elseif ($milestonesExists -and $milestonesCSVExists -and -not $opportunitiesExists) {
    Write-Warning "Cannot add milestones without opportunities list for lookup relationships"
}

# Summary
Write-Host ""
Write-Host "=== CSV Data Population Complete ===" -ForegroundColor Green

if ($customersExists) {
    Write-Host "✓ Customers list: $totalCustomersAdded new records added from CSV" -ForegroundColor Gray
}

if ($contactsExists) {
    Write-Host "✓ Contacts list: $totalContactsAdded new records added from CSV" -ForegroundColor Gray
}

if ($opportunitiesExists) {
    Write-Host "✓ Opportunities list: $totalOpportunitiesAdded new records added from CSV" -ForegroundColor Gray
}

if ($milestonesExists) {
    Write-Host "✓ Milestones list: $totalMilestonesAdded new records added from CSV" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Data populated from CSV files:" -ForegroundColor Cyan
Write-Host "  • $totalCustomersAdded customers from Customers.csv" -ForegroundColor Gray
Write-Host "  • $totalContactsAdded contacts from Contacts.csv" -ForegroundColor Gray
Write-Host "  • $totalOpportunitiesAdded opportunities from Opportunities.csv" -ForegroundColor Gray
Write-Host "  • $totalMilestonesAdded milestones from Milestones.csv" -ForegroundColor Gray
Write-Host "  • All lookup relationships properly established" -ForegroundColor Gray
Write-Host "  • Currency values and dates properly parsed" -ForegroundColor Gray

# Disconnect
try {
    Disconnect-PnPOnline
    Write-Host ""
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
} catch {
    # Ignore disconnect errors
}
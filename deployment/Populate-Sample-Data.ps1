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
Write-Host "=== SharePoint Enhanced Sample Data Population ===" -ForegroundColor Cyan
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
$contactsListName = "$($ListPrefix)Contacts"
$milestonesListName = "$($ListPrefix)Milestones"

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

# Function to generate normal distribution values
function Get-NormalDistributionValue {
    param(
        [double]$Min = 25000,
        [double]$Max = 2500000
    )
    
    # Simplified approach using weighted random values to approximate normal distribution
    $random1 = Get-Random -Minimum 0.0 -Maximum 1.0
    $random2 = Get-Random -Minimum 0.0 -Maximum 1.0
    $random3 = Get-Random -Minimum 0.0 -Maximum 1.0
    
    # Average of three random numbers approximates normal distribution
    $normalized = ($random1 + $random2 + $random3) / 3.0
    
    # Scale to our range
    return [Math]::Round($Min + ($Max - $Min) * $normalized, 0)
}

# Function to get random weekday within specified days
function Get-RandomWeekday {
    param(
        [int]$DaysFromNow = 512
    )
    
    do {
        # Get random number of days from now (1 to 512)
        $randomDays = Get-Random -Minimum 1 -Maximum 513
        $targetDate = (Get-Date).AddDays($randomDays)
        
        # Check if it's a weekday (Monday=1, Friday=5)
        $dayOfWeek = [int]$targetDate.DayOfWeek
        
    } while ($dayOfWeek -eq 0 -or $dayOfWeek -eq 6)  # 0=Sunday, 6=Saturday
    
    return $targetDate
}

# Function to get random recurring revenue amount based on total amount
function Get-RecurringRevenue {
    param(
        [double]$TotalAmount
    )
    
    # Generate recurring revenue as 5-25% of total amount
    $percentage = (Get-Random -Minimum 5 -Maximum 26) / 100.0
    return [Math]::Round($TotalAmount * $percentage, 0)
}

# Function to get random recurrence period - updated for the actual field type
function Get-RecurrencePeriod {
    # Based on the XML, RecurringRevenueModel is a Choice field, and Recurrences is a Number field
    # So Recurrences should be a number, not a text choice
    return Get-Random -Minimum 1 -Maximum 60  # 1 to 60 recurrences
}

# Function to get random recurring revenue model - this is the Choice field
function Get-RecurringRevenueModel {
    $modelOptions = @("Monthly", "Quarterly", "Semi-Annually", "Annually", "One-time", "Usage-based")
    return $modelOptions | Get-Random
}

# Function to get start date (5-90 days after close date)
function Get-StartDate {
    param(
        [DateTime]$CloseDate
    )
    
    $daysAfterClose = Get-Random -Minimum 5 -Maximum 91
    return $CloseDate.AddDays($daysAfterClose)
}

# Function to generate phone numbers
function Get-RandomPhoneNumber {
    $areaCode = Get-Random -Minimum 200 -Maximum 999
    $exchange = Get-Random -Minimum 200 -Maximum 999
    $number = Get-Random -Minimum 1000 -Maximum 9999
    return "$areaCode-$exchange-$number"
}

# Function to generate email from name
function Get-EmailFromName {
    param(
        [string]$FirstName,
        [string]$LastName,
        [string]$Domain
    )
    
    $firstName = $FirstName.ToLower().Replace(" ", "")
    $lastName = $LastName.ToLower().Replace(" ", "")
    return "$firstName.$lastName@$Domain"
}

# Function to add sample customers (expanded to 30)
function Add-SampleCustomers {
    param([string]$ListName)
    
    Write-Host "Adding sample customers to $ListName..." -ForegroundColor Yellow
    
    $sampleCustomers = @(
        @{
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
        },
        @{
            CustomerName = "Financial Dynamics"
            PrimaryContact = "Peter Jackson"
            PrimaryContactTitle = "President"
            AlternateContact = "Mary Wilson"
            AlternateContactTitle = "VP Operations"
            AlternateContact2 = "Tom Harris"
            AlternateContact2Title = "CFO"
            Website = "https://financialdynamics.com"
            NAICSSector = "522110"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Retail Excellence"
            PrimaryContact = "Amanda Clark"
            PrimaryContactTitle = "Store Manager"
            AlternateContact = "Kevin Martinez"
            AlternateContactTitle = "Regional Director"
            AlternateContact2 = "Nicole Brown"
            AlternateContact2Title = "Merchandising Manager"
            Website = "https://retailexcellence.com"
            NAICSSector = "452210"
            CustomerStatus = "Prospect"
        },
        @{
            CustomerName = "Energy Solutions Group"
            PrimaryContact = "Michael Turner"
            PrimaryContactTitle = "Engineering Director"
            AlternateContact = "Rachel Green"
            AlternateContactTitle = "Project Manager"
            AlternateContact2 = "Daniel White"
            AlternateContact2Title = "Safety Coordinator"
            Website = "https://energysolutions.com"
            NAICSSector = "221112"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Construction Dynamics"
            PrimaryContact = "Steve Robinson"
            PrimaryContactTitle = "General Manager"
            AlternateContact = "Linda Moore"
            AlternateContactTitle = "Office Manager"
            AlternateContact2 = "Gary Thompson"
            AlternateContact2Title = "Site Supervisor"
            Website = "https://constructiondynamics.com"
            NAICSSector = "236220"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Education First"
            PrimaryContact = "Dr. Jennifer Adams"
            PrimaryContactTitle = "Superintendent"
            AlternateContact = "Robert Lewis"
            AlternateContactTitle = "IT Director"
            AlternateContact2 = "Susan Walker"
            AlternateContact2Title = "Curriculum Director"
            Website = "https://educationfirst.com"
            NAICSSector = "611110"
            CustomerStatus = "Prospect"
        },
        @{
            CustomerName = "Transport Logistics"
            PrimaryContact = "Carlos Rodriguez"
            PrimaryContactTitle = "Fleet Manager"
            AlternateContact = "Angela Davis"
            AlternateContactTitle = "Operations Manager"
            AlternateContact2 = "Bryan Miller"
            AlternateContact2Title = "Maintenance Supervisor"
            Website = "https://transportlogistics.com"
            NAICSSector = "484110"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Food Service Partners"
            PrimaryContact = "Maria Gonzalez"
            PrimaryContactTitle = "General Manager"
            AlternateContact = "John Wilson"
            AlternateContactTitle = "Kitchen Manager"
            AlternateContact2 = "Lisa Garcia"
            AlternateContact2Title = "Service Manager"
            Website = "https://foodservicepartners.com"
            NAICSSector = "722511"
            CustomerStatus = "Prospect"
        },
        @{
            CustomerName = "Professional Services Group"
            PrimaryContact = "Thomas Anderson"
            PrimaryContactTitle = "Managing Partner"
            AlternateContact = "Patricia Taylor"
            AlternateContactTitle = "Senior Associate"
            AlternateContact2 = "Richard Brown"
            AlternateContact2Title = "Operations Manager"
            Website = "https://professionalservices.com"
            NAICSSector = "541110"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Media Communications"
            PrimaryContact = "Jessica Williams"
            PrimaryContactTitle = "Creative Director"
            AlternateContact = "Andrew Johnson"
            AlternateContactTitle = "Production Manager"
            AlternateContact2 = "Michelle Davis"
            AlternateContact2Title = "Account Manager"
            Website = "https://mediacommunications.com"
            NAICSSector = "541810"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Pharmaceutical Research"
            PrimaryContact = "Dr. William Chen"
            PrimaryContactTitle = "Research Director"
            AlternateContact = "Karen Lee"
            AlternateContactTitle = "Lab Manager"
            AlternateContact2 = "David Kim"
            AlternateContact2Title = "Regulatory Affairs"
            Website = "https://pharmresearch.com"
            NAICSSector = "325412"
            CustomerStatus = "Prospect"
        },
        @{
            CustomerName = "Real Estate Ventures"
            PrimaryContact = "Catherine Moore"
            PrimaryContactTitle = "Broker"
            AlternateContact = "James Parker"
            AlternateContactTitle = "Property Manager"
            AlternateContact2 = "Sandra Clark"
            AlternateContact2Title = "Development Manager"
            Website = "https://realestateventures.com"
            NAICSSector = "531210"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Insurance Specialists"
            PrimaryContact = "Robert Martinez"
            PrimaryContactTitle = "Agency Owner"
            AlternateContact = "Nancy Thompson"
            AlternateContactTitle = "Claims Manager"
            AlternateContact2 = "Paul Anderson"
            AlternateContact2Title = "Underwriter"
            Website = "https://insurancespecialists.com"
            NAICSSector = "524210"
            CustomerStatus = "Prospect"
        },
        @{
            CustomerName = "Agricultural Systems"
            PrimaryContact = "Frank Wilson"
            PrimaryContactTitle = "Farm Manager"
            AlternateContact = "Helen Garcia"
            AlternateContactTitle = "Crop Specialist"
            AlternateContact2 = "Mark Rodriguez"
            AlternateContact2Title = "Equipment Manager"
            Website = "https://agriculturalsystems.com"
            NAICSSector = "111110"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Hotel Management Group"
            PrimaryContact = "Diana Adams"
            PrimaryContactTitle = "General Manager"
            AlternateContact = "Christopher Lee"
            AlternateContactTitle = "Front Desk Manager"
            AlternateContact2 = "Maria Santos"
            AlternateContact2Title = "Housekeeping Manager"
            Website = "https://hotelmanagement.com"
            NAICSSector = "721110"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Automotive Parts Plus"
            PrimaryContact = "Gregory Johnson"
            PrimaryContactTitle = "Store Manager"
            AlternateContact = "Brenda Davis"
            AlternateContactTitle = "Parts Specialist"
            AlternateContact2 = "Tony Martinez"
            AlternateContact2Title = "Service Advisor"
            Website = "https://autopartsplus.com"
            NAICSSector = "441310"
            CustomerStatus = "Prospect"
        },
        @{
            CustomerName = "Textile Manufacturing"
            PrimaryContact = "Victoria Brown"
            PrimaryContactTitle = "Production Manager"
            AlternateContact = "Samuel Wilson"
            AlternateContactTitle = "Quality Control"
            AlternateContact2 = "Rebecca Taylor"
            AlternateContact2Title = "Logistics Coordinator"
            Website = "https://textilemanufacturing.com"
            NAICSSector = "313210"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Environmental Services"
            PrimaryContact = "Jonathan Green"
            PrimaryContactTitle = "Environmental Engineer"
            AlternateContact = "Laura Miller"
            AlternateContactTitle = "Project Coordinator"
            AlternateContact2 = "Kevin White"
            AlternateContact2Title = "Field Technician"
            Website = "https://environmentalservices.com"
            NAICSSector = "562910"
            CustomerStatus = "Prospect"
        },
        @{
            CustomerName = "Sports & Recreation"
            PrimaryContact = "Marcus Johnson"
            PrimaryContactTitle = "Facility Manager"
            AlternateContact = "Stephanie Clark"
            AlternateContactTitle = "Program Director"
            AlternateContact2 = "Derek Anderson"
            AlternateContact2Title = "Maintenance Manager"
            Website = "https://sportsrecreation.com"
            NAICSSector = "713940"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Security Solutions"
            PrimaryContact = "Raymond Davis"
            PrimaryContactTitle = "Security Director"
            AlternateContact = "Carol Thompson"
            AlternateContactTitle = "Operations Manager"
            AlternateContact2 = "Eugene Garcia"
            AlternateContact2Title = "Technical Specialist"
            Website = "https://securitysolutions.com"
            NAICSSector = "561612"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Chemical Processing"
            PrimaryContact = "Donna Martinez"
            PrimaryContactTitle = "Plant Manager"
            AlternateContact = "Albert Rodriguez"
            AlternateContactTitle = "Process Engineer"
            AlternateContact2 = "Ruth Wilson"
            AlternateContact2Title = "Safety Manager"
            Website = "https://chemicalprocessing.com"
            NAICSSector = "325180"
            CustomerStatus = "Prospect"
        },
        @{
            CustomerName = "Publishing House"
            PrimaryContact = "Arthur Lee"
            PrimaryContactTitle = "Editor-in-Chief"
            AlternateContact = "Evelyn Brown"
            AlternateContactTitle = "Production Manager"
            AlternateContact2 = "Walter Kim"
            AlternateContact2Title = "Marketing Director"
            Website = "https://publishinghouse.com"
            NAICSSector = "511130"
            CustomerStatus = "Prospect"
        },
        @{
            CustomerName = "Mining Operations"
            PrimaryContact = "Harold Johnson"
            PrimaryContactTitle = "Site Manager"
            AlternateContact = "Gloria Davis"
            AlternateContactTitle = "Safety Coordinator"
            AlternateContact2 = "Ralph Martinez"
            AlternateContact2Title = "Equipment Supervisor"
            Website = "https://miningoperations.com"
            NAICSSector = "212230"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Telecommunications Inc"
            PrimaryContact = "Betty Wilson"
            PrimaryContactTitle = "Network Manager"
            AlternateContact = "Ernest Thompson"
            AlternateContactTitle = "Technical Support"
            AlternateContact2 = "Mildred Garcia"
            AlternateContact2Title = "Customer Service Manager"
            Website = "https://telecommunications.com"
            NAICSSector = "517110"
            CustomerStatus = "Active"
        },
        @{
            CustomerName = "Waste Management Co"
            PrimaryContact = "Gerald Rodriguez"
            PrimaryContactTitle = "Operations Director"
            AlternateContact = "Joan Anderson"
            AlternateContactTitle = "Route Manager"
            AlternateContact2 = "Frank Miller"
            AlternateContact2Title = "Disposal Coordinator"
            Website = "https://wastemanagement.com"
            NAICSSector = "562111"
            CustomerStatus = "Prospect"
        },
        @{
            CustomerName = "Consulting Experts"
            PrimaryContact = "Louise Clark"
            PrimaryContactTitle = "Senior Consultant"
            AlternateContact = "Wayne Taylor"
            AlternateContactTitle = "Project Manager"
            AlternateContact2 = "Theresa Brown"
            AlternateContact2Title = "Business Analyst"
            Website = "https://consultingexperts.com"
            NAICSSector = "541611"
            CustomerStatus = "Active"
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

# Function to add sample contacts (2-3 per customer)
function Add-SampleContacts {
    param(
        [string]$ListName,
        [string]$CustomersListName
    )
    
    Write-Host "Adding sample contacts to $ListName..." -ForegroundColor Yellow
    
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
    
    # Sample contact data templates
    $firstNames = @("John", "Jane", "Michael", "Sarah", "David", "Lisa", "Robert", "Jennifer", "William", "Mary", 
                   "James", "Patricia", "Richard", "Linda", "Charles", "Barbara", "Joseph", "Elizabeth", "Thomas", "Maria",
                   "Christopher", "Susan", "Daniel", "Margaret", "Matthew", "Dorothy", "Anthony", "Nancy", "Mark", "Karen",
                   "Donald", "Helen", "Steven", "Sandra", "Paul", "Donna", "Andrew", "Carol", "Joshua", "Ruth")
    
    $lastNames = @("Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez",
                  "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin",
                  "Lee", "Perez", "Thompson", "White", "Harris", "Sanchez", "Clark", "Ramirez", "Lewis", "Robinson",
                  "Walker", "Young", "Allen", "King", "Wright", "Scott", "Torres", "Nguyen", "Hill", "Flores")
    
    $jobTitles = @("Manager", "Director", "Coordinator", "Specialist", "Analyst", "Administrator", "Supervisor", 
                  "Executive", "Officer", "Assistant", "Representative", "Consultant", "Engineer", "Technician",
                  "VP", "President", "CEO", "CFO", "CTO", "COO", "Senior Manager", "Project Manager", "Operations Manager",
                  "Sales Manager", "Marketing Manager", "IT Manager", "HR Manager", "Finance Manager", "Quality Manager")
    
    $addedCount = 0
    
    foreach ($customer in $customers) {
        try {
            # Get customer domain from website URL if available
            $customerDomain = "company.com"  # Default
            if ($customer["Website"]) {
                try {
                    $uri = [System.Uri]$customer["Website"]
                    $customerDomain = $uri.Host.ToLower()
                } catch {
                    # Use customer name as domain if URL parsing fails
                    $cleanName = $customer["CustomerName"].Replace(" ", "").Replace("&", "").Replace(".", "").ToLower()
                    $customerDomain = "$cleanName.com"
                }
            } else {
                # Use customer name as domain
                $cleanName = $customer["CustomerName"].Replace(" ", "").Replace("&", "").Replace(".", "").ToLower()
                $customerDomain = "$cleanName.com"
            }
            
            # Generate 2-3 contacts per customer
            $numContacts = Get-Random -Minimum 2 -Maximum 4
            
            for ($i = 0; $i -lt $numContacts; $i++) {
                $firstName = $firstNames | Get-Random
                $lastName = $lastNames | Get-Random
                $fullName = "$firstName $lastName"
                $jobTitle = $jobTitles | Get-Random
                $email = Get-EmailFromName -FirstName $firstName -LastName $lastName -Domain $customerDomain
                $officePhone = Get-RandomPhoneNumber
                $mobilePhone = Get-RandomPhoneNumber
                
                $contact = @{
                    ContactName = $fullName
                    JobTitle = $jobTitle
                    OfficePhone = $officePhone
                    MobilePhone = $mobilePhone
                    Email = $email
                }
                
                # Add customer lookup
                $contact[$customerFieldName] = $customer.Id
                
                # Check if contact already exists (by name and customer)
                $existingContact = Get-PnPListItem -List $ListName -Query "<View><Query><Where><And><Eq><FieldRef Name='ContactName'/><Value Type='Text'>$fullName</Value></Eq><Eq><FieldRef Name='$customerFieldName'/><Value Type='Lookup'>$($customer.Id)</Value></Eq></And></Where></Query></View>" -ErrorAction SilentlyContinue
                
                if (-not $existingContact) {
                    Add-PnPListItem -List $ListName -Values $contact -ErrorAction Stop
                    Write-Host "  ✓ Added contact: $fullName ($jobTitle) at $($customer['CustomerName'])" -ForegroundColor Green
                    $addedCount++
                } else {
                    Write-Host "  - Contact already exists: $fullName at $($customer['CustomerName'])" -ForegroundColor Gray
                }
            }
            
        } catch {
            Write-Warning "  Failed to add contacts for customer $($customer['CustomerName']): $($_.Exception.Message)"
        }
    }
    
    Write-Host "  ✓ Added $addedCount new contacts" -ForegroundColor Green
    return $addedCount
}

# Function to add sample opportunities (expanded to 40)
function Add-SampleOpportunities {
    param(
        [string]$ListName,
        [string]$CustomersListName
    )
    
    Write-Host "Adding sample opportunities to $ListName..." -ForegroundColor Yellow
    
    # Wait a moment for any recent field creation to propagate
    Write-Host "  Waiting for field propagation..." -ForegroundColor Gray
    Start-Sleep -Seconds 3
    
    # Check what fields exist in the list
    Write-Host "  Analyzing list fields..." -ForegroundColor Gray
    $listFields = Get-PnPField -List $ListName
    $availableFields = @{}
    foreach ($field in $listFields) {
        $availableFields[$field.InternalName] = $field.Title
    }
    
    # Debug: Show all fields in the list
    Write-Host "  Available fields in list:" -ForegroundColor Gray
    foreach ($fieldName in ($availableFields.Keys | Sort-Object)) {
        $fieldTitle = $availableFields[$fieldName]
        $field = $listFields | Where-Object { $_.InternalName -eq $fieldName }
        $fieldType = if ($field) { $field.TypeDisplayName } else { "Unknown" }
        Write-Host "    $fieldName -> $fieldTitle ($fieldType)" -ForegroundColor Gray
    }
    Write-Host ""
    
    # Check for recurring revenue fields with multiple possible names
    $recurringRevenueField = $null
    $recurrenceField = $null
    $startDateField = $null
    
    # Try different possible field names for RecurringRevenue - put the correct one first
    $possibleRecurringNames = @("RecurringRevenueModel", "RecurringRevenue", "Recurring_x0020_Revenue", "RecurringRevenue0", "Recurring", "RecurringAmount")
    foreach ($fieldName in $possibleRecurringNames) {
        if ($availableFields.ContainsKey($fieldName)) {
            $recurringRevenueField = $fieldName
            break
        }
    }
    
    # Try different possible field names for Recurrence - Recurrences is correct
    $possibleRecurrenceNames = @("Recurrences", "Recurrence", "Recurrence0", "RecurrencePeriod", "Recurrent", "RecurringPeriod")
    foreach ($fieldName in $possibleRecurrenceNames) {
        if ($availableFields.ContainsKey($fieldName)) {
            $recurrenceField = $fieldName
            break
        }
    }
    
    # Try different possible field names for StartDate
    $possibleStartDateNames = @("StartDate", "Start_x0020_Date", "StartDate0", "Start")
    foreach ($fieldName in $possibleStartDateNames) {
        if ($availableFields.ContainsKey($fieldName)) {
            $startDateField = $fieldName
            break
        }
    }
    
    Write-Host "  Field mapping results:" -ForegroundColor Gray
    Write-Host "    RecurringRevenue: $(if ($recurringRevenueField) { "✓ $recurringRevenueField" } else { '✗ Not found' })" -ForegroundColor $(if ($recurringRevenueField) { 'Green' } else { 'Red' })
    Write-Host "    Recurrence: $(if ($recurrenceField) { "✓ $recurrenceField" } else { '✗ Not found' })" -ForegroundColor $(if ($recurrenceField) { 'Green' } else { 'Red' })
    Write-Host "    StartDate: $(if ($startDateField) { "✓ $startDateField" } else { '✗ Not found' })" -ForegroundColor $(if ($startDateField) { 'Green' } else { 'Red' })
    
    # Check what the actual CustomerID lookup field internal name is
    $customerFieldName = $null
    $possibleCustomerFieldNames = @("CustomerName", "CustomerIDId", "CustomerId", "CustomerID", "CustomerIdId", "CustomerIDLookupId", "Customer", "CustomerLookup")
    
    foreach ($fieldName in $possibleCustomerFieldNames) {
        if ($availableFields.ContainsKey($fieldName)) {
            # Double-check it's actually a lookup field
            $field = $listFields | Where-Object { $_.InternalName -eq $fieldName }
            if ($field -and $field.TypeDisplayName -eq "Lookup") {
                $customerFieldName = $fieldName
                Write-Host "  Found customer lookup field: $fieldName (Display: $($availableFields[$fieldName]))" -ForegroundColor Gray
                break
            }
        }
    }
    
    if (-not $customerFieldName) {
        Write-Warning "CustomerID lookup field not found in $ListName. Trying to find customer-related lookup fields..."
        foreach ($fieldName in $availableFields.Keys) {
            $field = $listFields | Where-Object { $_.InternalName -eq $fieldName }
            if ($field -and $field.TypeDisplayName -eq "Lookup" -and ($fieldName -like "*Customer*" -or $availableFields[$fieldName] -like "*Customer*")) {
                $customerFieldName = $fieldName
                Write-Host "  Using customer lookup field: $fieldName (Display: $($availableFields[$fieldName]))" -ForegroundColor Yellow
                break
            }
        }
        
        if (-not $customerFieldName) {
            Write-Warning "No customer lookup fields found. Please ensure a customer lookup field exists in the opportunities list."
            Write-Host "  Available lookup fields:" -ForegroundColor Gray
            foreach ($field in ($listFields | Where-Object { $_.TypeDisplayName -eq "Lookup" })) {
                Write-Host "    $($field.InternalName) -> $($field.Title)" -ForegroundColor Gray
            }
            return 0
        }
    }
    
    Write-Host "  Customer field mapping: $customerFieldName" -ForegroundColor Green
    
    # Get customer items for lookup
    try {
        $customers = Get-PnPListItem -List $CustomersListName -ErrorAction Stop
        if ($customers.Count -eq 0) {
            Write-Warning "No customers found in $CustomersListName. Cannot create opportunities."
            return 0
        }
        Write-Host "  Found $($customers.Count) customers for lookup" -ForegroundColor Gray
    } catch {
        Write-Error "Failed to get customers from $CustomersListName : $($_.Exception.Message)"
        return 0
    }
    
    # Define opportunity templates with calculated distributions
    $opportunityTemplates = @(
        "ERP Implementation Project", "Cloud Migration Initiative", "Security Assessment & Audit", 
        "Manufacturing Automation System", "Healthcare Data Analytics Platform", "Digital Transformation Project",
        "Network Infrastructure Upgrade", "Business Intelligence Solution", "Customer Portal Development",
        "Mobile Application Development", "Data Center Consolidation", "Cybersecurity Enhancement",
        "E-commerce Platform", "Document Management System", "Training & Development Program",
        "Quality Management System", "Supply Chain Optimization", "Financial Planning Software",
        "HR Management System", "Marketing Automation Platform", "CRM System Upgrade",
        "Inventory Management Solution", "Point of Sale System", "Compliance Management Tool",
        "Project Management Software", "Video Conferencing Solution", "Backup & Recovery System",
        "Website Redesign Project", "Social Media Management Tool", "Customer Service Platform",
        "Fleet Management System", "Environmental Monitoring Solution", "Safety Management System",
        "Performance Analytics Dashboard", "Workflow Automation Tool", "Integration Services Project",
        "Database Migration Project", "Legacy System Modernization", "API Development Project",
        "Machine Learning Implementation"
    )
    
    $stages = @("Lead Qualification", "Nurturing", "Proposal", "Negotiation", "Project Execution", "Closeout")
    $owners = @("John Manager", "Sarah Director", "Mike Consultant", "Lisa Engineer", "Dr. Tech Advisor", 
                "Amanda Sales", "Robert Analyst", "Jennifer PM", "David Specialist", "Maria Coordinator")
    $milestones = @("Technical Review", "Contract Finalization", "Stakeholder Meeting", "Site Survey", 
                   "ROI Presentation", "Proof of Concept", "Requirements Analysis", "Proposal Review", 
                   "Budget Approval", "Implementation Planning")
    
    # Create stage distribution array to ensure all stages are represented
    $stageAssignments = @()
    
    # Calculate opportunities per stage (40 opportunities / 6 stages = 6 per stage, with 4 remainder)
    $opportunitiesPerStage = [Math]::Floor(40 / $stages.Count)
    $remainingOpportunities = 40 - ($opportunitiesPerStage * $stages.Count)
    
    # Assign stages evenly (6 opportunities per stage)
    for ($stageIndex = 0; $stageIndex -lt $stages.Count; $stageIndex++) {
        for ($i = 0; $i -lt $opportunitiesPerStage; $i++) {
            $stageAssignments += $stages[$stageIndex]
        }
    }
    
    # Distribute the remaining 4 opportunities across the first 4 stages
    for ($i = 0; $i -lt $remainingOpportunities; $i++) {
        $stageAssignments += $stages[$i]
    }
    
    # Shuffle the stage assignments to randomize
    $stageAssignments = $stageAssignments | Get-Random -Count $stageAssignments.Count
    
    Write-Host "  Stage distribution: 6-7 opportunities across 6 stages (Lead Qualification, Nurturing, Proposal, Negotiation, Project Execution, Closeout)" -ForegroundColor Gray
    
    # Create customer assignment array according to specifications
    $customerAssignments = @()
    
    # Ensure we have enough customers
    if ($customers.Count -lt 8) {
        Write-Warning "Not enough customers found. Need at least 8 customers for proper distribution."
        # Fall back to simple round-robin assignment
        for ($i = 0; $i -lt 40; $i++) {
            $customerAssignments += ($i % $customers.Count)
        }
    } else {
        # Assign 5 opportunities to customer 0 (index 0)
        for ($i = 0; $i -lt 5; $i++) { $customerAssignments += 0 }
        
        # Assign 3 opportunities each to customers 1 and 2 (indices 1, 2)
        for ($i = 0; $i -lt 3; $i++) { $customerAssignments += 1 }
        for ($i = 0; $i -lt 3; $i++) { $customerAssignments += 2 }
        
        # Skip customers 3, 4, 5, 6, 7 (5 customers with no opportunities)
        
        # Distribute remaining 29 opportunities among remaining customers (indices 8 and up)
        $remainingCustomerCount = $customers.Count - 8
        if ($remainingCustomerCount -gt 0) {
            for ($i = 0; $i -lt 29; $i++) {
                $customerIndex = 8 + ($i % $remainingCustomerCount)
                $customerAssignments += $customerIndex
            }
        } else {
            # If we don't have enough customers, distribute among available ones
            for ($i = 0; $i -lt 29; $i++) {
                $customerAssignments += ($i % $customers.Count)
            }
        }
        
        # Shuffle the assignments to randomize
        $customerAssignments = $customerAssignments | Get-Random -Count $customerAssignments.Count
    }
    
    Write-Host "  Customer assignment distribution prepared" -ForegroundColor Gray
    
    $addedCount = 0
    
    for ($i = 0; $i -lt 40; $i++) {
        try {
            Write-Host "  Processing opportunity $($i + 1) of 40..." -ForegroundColor Gray
            
            # Generate opportunity values with normal distribution
            $amount = Get-NormalDistributionValue -Min 25000 -Max 2500000
            
            # Assign probability based on distribution: 30% Low, 50% Medium, 20% High
            $probRand = Get-Random -Minimum 1 -Maximum 101
            if ($probRand -le 30) { $probability = "Low" }
            elseif ($probRand -le 80) { $probability = "Medium" }
            else { $probability = "High" }
            
            # Assign status based on distribution: 5% Critical, 15% At Risk, 80% Active
            $statusRand = Get-Random -Minimum 1 -Maximum 101
            if ($statusRand -le 5) { $status = "Critical" }
            elseif ($statusRand -le 20) { $status = "At Risk" }
            else { $status = "Active" }
            
            # Generate close date
            $closeDate = Get-RandomWeekday -DaysFromNow 512
            
            # Create base opportunity object
            $opportunity = @{
                OpportunityName = $opportunityTemplates[$i]
                Status = $status
                OpportunityStage = $stageAssignments[$i]
                Amount = $amount
                Probability = $probability
                OpportunityOwner = $owners | Get-Random
                Close = $closeDate
                NextMilestone = $milestones | Get-Random
                NextMilestoneDate = (Get-Date).AddDays((Get-Random -Minimum 5 -Maximum 60))
            }
            
            # Add recurring revenue fields using the correct field names and types
            if ($recurringRevenueField) {
                # RecurringRevenueModel is a Choice field - use the model function
                $recurringRevenueModel = Get-RecurringRevenueModel
                $opportunity[$recurringRevenueField] = $recurringRevenueModel
            }
            
            if ($recurrenceField) {
                # Recurrences is a Number field - use random number
                $recurrences = Get-RecurrencePeriod  # This now returns a number
                $opportunity[$recurrenceField] = $recurrences
            }
            
            if ($startDateField) {
                $startDate = Get-StartDate -CloseDate $closeDate
                $opportunity[$startDateField] = $startDate
            }
            
            # Assign customer based on our distribution plan
            $customerIndex = $customerAssignments[$i]
            if ($customerIndex -ge $customers.Count) {
                $customerIndex = $i % $customers.Count
            }
            $customer = $customers[$customerIndex]
            
            # Debug: Show what customer field we're trying to use
            Write-Host "    Using customer field: $customerFieldName with customer ID: $($customer.Id)" -ForegroundColor Gray
            
            $opportunity[$customerFieldName] = $customer.Id
            
            # Check if opportunity already exists
            $existingOpp = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='OpportunityName'/><Value Type='Text'>$($opportunity.OpportunityName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
            
            if (-not $existingOpp) {
                Add-PnPListItem -List $ListName -Values $opportunity -ErrorAction Stop
                
                # Build display message based on available fields
                $displayMessage = "  ✓ Added opportunity: $($opportunity.OpportunityName) ($($opportunity.Amount.ToString('C0')))"
                if ($recurringRevenueField -and $opportunity.ContainsKey($recurringRevenueField)) {
                    $displayMessage += " - Model: $($opportunity[$recurringRevenueField])"
                }
                if ($recurrenceField -and $opportunity.ContainsKey($recurrenceField)) {
                    $displayMessage += " ($($opportunity[$recurrenceField]) times)"
                }
                
                Write-Host $displayMessage -ForegroundColor Green
                $addedCount++
            } else {
                Write-Host "  - Opportunity already exists: $($opportunity.OpportunityName)" -ForegroundColor Gray
            }
        } catch {
            Write-Warning "  Failed to add opportunity $($opportunityTemplates[$i]): $($_.Exception.Message)"
        }
    }
    
    Write-Host "  ✓ Added $addedCount new opportunities" -ForegroundColor Green
    return $addedCount
}

# Function to add sample milestones (3-5 per opportunity)
function Add-SampleMilestones {
    param(
        [string]$ListName,
        [string]$OpportunitiesListName
    )
    
    Write-Host "Adding sample milestones to $ListName..." -ForegroundColor Yellow
    
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
    
    # Milestone templates organized by opportunity stage
    $milestoneTemplates = @{
        "Lead Qualification" = @(
            @{Name="Initial Contact"; Days=0; Status="Completed"},
            @{Name="Needs Assessment"; Days=7; Status="In Progress"},
            @{Name="Stakeholder Identification"; Days=14; Status="Not Started"},
            @{Name="Budget Confirmation"; Days=21; Status="Not Started"}
        )
        "Nurturing" = @(
            @{Name="Relationship Building"; Days=0; Status="In Progress"},
            @{Name="Pain Point Analysis"; Days=10; Status="In Progress"},
            @{Name="Solution Mapping"; Days=20; Status="Not Started"},
            @{Name="Competitive Analysis"; Days=30; Status="Not Started"},
            @{Name="ROI Calculation"; Days=40; Status="Not Started"}
        )
        "Proposal" = @(
            @{Name="Requirements Gathering"; Days=0; Status="Completed"},
            @{Name="Technical Specification"; Days=7; Status="In Progress"},
            @{Name="Proposal Development"; Days=14; Status="In Progress"},
            @{Name="Pricing Analysis"; Days=21; Status="Not Started"},
            @{Name="Proposal Review"; Days=28; Status="Not Started"}
        )
        "Negotiation" = @(
            @{Name="Contract Terms Review"; Days=0; Status="In Progress"},
            @{Name="Legal Review"; Days=7; Status="In Progress"},
            @{Name="Pricing Negotiation"; Days=14; Status="Not Started"},
            @{Name="Final Terms Agreement"; Days=21; Status="Not Started"}
        )
        "Project Execution" = @(
            @{Name="Project Kickoff"; Days=0; Status="Completed"},
            @{Name="Phase 1 Delivery"; Days=30; Status="In Progress"},
            @{Name="Phase 2 Delivery"; Days=60; Status="Not Started"},
            @{Name="Testing & QA"; Days=90; Status="Not Started"},
            @{Name="User Training"; Days=100; Status="Not Started"}
        )
        "Closeout" = @(
            @{Name="Final Delivery"; Days=0; Status="Completed"},
            @{Name="Customer Acceptance"; Days=7; Status="In Progress"},
            @{Name="Documentation Handover"; Days=14; Status="Not Started"},
            @{Name="Post-Implementation Review"; Days=30; Status="Not Started"}
        )
    }
    
    $owners = @("John Manager", "Sarah Director", "Mike Consultant", "Lisa Engineer", "Dr. Tech Advisor", 
                "Amanda Sales", "Robert Analyst", "Jennifer PM", "David Specialist", "Maria Coordinator",
                "Project Lead", "Technical Lead", "Business Analyst", "Quality Assurance", "Implementation Specialist")
    
    $addedCount = 0
    
    foreach ($opportunity in $opportunities) {
        try {
            # Get opportunity stage to determine appropriate milestones
            $oppStage = $opportunity["OpportunityStage"]
            if (-not $oppStage -or -not $milestoneTemplates.ContainsKey($oppStage)) {
                Write-Warning "  Unknown opportunity stage '$oppStage' for opportunity $($opportunity['OpportunityName']). Using default milestones."
                $oppStage = "Proposal"  # Default fallback
            }
            
            # Get milestone templates for this stage
            $stageTemplates = $milestoneTemplates[$oppStage]
            
            # Get opportunity close date to calculate milestone dates
            $oppCloseDate = $opportunity["Close"]
            if (-not $oppCloseDate) {
                $oppCloseDate = (Get-Date).AddDays(60)  # Default if no close date
            }
            
            # Create milestones for this opportunity
            foreach ($template in $stageTemplates) {
                try {
                    # Calculate milestone date based on template offset from close date
                    $milestoneDate = $oppCloseDate.AddDays(-$template.Days)
                    
                    # Ensure milestone date is not in the past (adjust if needed)
                    if ($milestoneDate -lt (Get-Date)) {
                        $milestoneDate = (Get-Date).AddDays((Get-Random -Minimum 1 -Maximum 30))
                    }
                    
                    $milestone = @{
                        MilestoneName = $template.Name
                        Owner = $owners | Get-Random
                        MilestoneDate = $milestoneDate
                        MilestoneStatus = $template.Status
                    }
                    
                    # Add opportunity lookup
                    $milestone[$opportunityFieldName] = $opportunity.Id
                    
                    # Check if milestone already exists for this opportunity
                    $existingMilestone = Get-PnPListItem -List $ListName -Query "<View><Query><Where><And><Eq><FieldRef Name='MilestoneName'/><Value Type='Text'>$($template.Name)</Value></Eq><Eq><FieldRef Name='$opportunityFieldName'/><Value Type='Lookup'>$($opportunity.Id)</Value></Eq></And></Where></Query></View>" -ErrorAction SilentlyContinue
                    
                    if (-not $existingMilestone) {
                        Add-PnPListItem -List $ListName -Values $milestone -ErrorAction Stop
                        Write-Host "  ✓ Added milestone: $($template.Name) for $($opportunity['OpportunityName'])" -ForegroundColor Green
                        $addedCount++
                    } else {
                        Write-Host "  - Milestone already exists: $($template.Name) for $($opportunity['OpportunityName'])" -ForegroundColor Gray
                    }
                    
                } catch {
                    Write-Warning "  Failed to add milestone $($template.Name) for opportunity $($opportunity['OpportunityName']): $($_.Exception.Message)"
                }
            }
            
        } catch {
            Write-Warning "  Failed to process milestones for opportunity $($opportunity['OpportunityName']): $($_.Exception.Message)"
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

if (-not $customersExists -and -not $opportunitiesExists -and -not $contactsExists -and -not $milestonesExists) {
    Write-Error "No lists found. Please check the list prefix and ensure lists have been created."
    Disconnect-PnPOnline
    exit 1
}

# Populate sample data
Write-Host ""
Write-Host "Populating sample data..." -ForegroundColor Yellow

$totalCustomersAdded = 0
$totalContactsAdded = 0
$totalOpportunitiesAdded = 0
$totalMilestonesAdded = 0

# Add customers first (required for other lookups)
if ($customersExists) {
    Write-Host ""
    $totalCustomersAdded = Add-SampleCustomers -ListName $customersListName
}

# Add contacts (requires customers to exist for lookups)
if ($contactsExists -and $customersExists) {
    Write-Host ""
    $totalContactsAdded = Add-SampleContacts -ListName $contactsListName -CustomersListName $customersListName
} elseif ($contactsExists -and -not $customersExists) {
    Write-Warning "Cannot add contacts without customers list for lookup relationships"
}

# Add opportunities (requires customers to exist for lookups)
if ($opportunitiesExists -and $customersExists) {
    Write-Host ""
    $totalOpportunitiesAdded = Add-SampleOpportunities -ListName $opportunitiesListName -CustomersListName $customersListName
} elseif ($opportunitiesExists -and -not $customersExists) {
    Write-Warning "Cannot add opportunities without customers list for lookup relationships"
}

# Add milestones (requires opportunities to exist for lookups)
if ($milestonesExists -and $opportunitiesExists) {
    Write-Host ""
    $totalMilestonesAdded = Add-SampleMilestones -ListName $milestonesListName -OpportunitiesListName $opportunitiesListName
} elseif ($milestonesExists -and -not $opportunitiesExists) {
    Write-Warning "Cannot add milestones without opportunities list for lookup relationships"
}

# Summary
Write-Host ""
Write-Host "=== Enhanced Sample Data Population Complete ===" -ForegroundColor Green

if ($customersExists) {
    Write-Host "✓ Customers list: $totalCustomersAdded new records added (30 total)" -ForegroundColor Gray
    Write-Host "  Distribution: Various industries and company types" -ForegroundColor Gray
}

if ($contactsExists) {
    Write-Host "✓ Contacts list: $totalContactsAdded new records added (60-90 total)" -ForegroundColor Gray
    Write-Host "  Distribution: 2-3 contacts per customer with realistic contact information" -ForegroundColor Gray
}

if ($opportunitiesExists) {
    Write-Host "✓ Opportunities list: $totalOpportunitiesAdded new records added (40 total)" -ForegroundColor Gray
    Write-Host "  Value range: $25,000 - $2,500,000 (normal distribution)" -ForegroundColor Gray
    Write-Host "  Probability: 30% Low, 50% Medium, 20% High" -ForegroundColor Gray
    Write-Host "  Status: 80% Active, 15% At Risk, 5% Critical" -ForegroundColor Gray
    Write-Host "  Customer distribution: 1 customer (5 opps), 2 customers (3 opps each)," -ForegroundColor Gray
    Write-Host "                         5 customers (0 opps), remaining distributed" -ForegroundColor Gray
    Write-Host "  Recurring revenue models: Monthly, Quarterly, Semi-Annually, Annually, One-time, Usage-based" -ForegroundColor Gray
    Write-Host "  Recurrence counts: 1-60 occurrences" -ForegroundColor Gray
    Write-Host "  Start dates: 5-90 days after close date" -ForegroundColor Gray
}

if ($milestonesExists) {
    Write-Host "✓ Milestones list: $totalMilestonesAdded new records added (160-200 total)" -ForegroundColor Gray
    Write-Host "  Distribution: 4-5 stage-appropriate milestones per opportunity" -ForegroundColor Gray
    Write-Host "  Status distribution: Mix of Completed, In Progress, and Not Started" -ForegroundColor Gray
    Write-Host "  Dates: Calculated based on opportunity close dates and stage requirements" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Enhanced sample data includes:" -ForegroundColor Cyan
Write-Host "  • 30 diverse customers across multiple industries" -ForegroundColor Gray
Write-Host "  • 60-90 contacts (2-3 per customer) with realistic contact details" -ForegroundColor Gray
Write-Host "  • 40 opportunities with realistic value distribution" -ForegroundColor Gray
Write-Host "  • 160-200 milestones (4-5 per opportunity) with stage-appropriate tasks" -ForegroundColor Gray
Write-Host "  • Normal distribution of opportunity values ($25K-$2.5M)" -ForegroundColor Gray
Write-Host "  • Varied opportunity stages and probability levels" -ForegroundColor Gray
Write-Host "  • Strategic customer-opportunity assignments" -ForegroundColor Gray
Write-Host "  • Realistic contact information and milestone dates" -ForegroundColor Gray
Write-Host "  • Proper lookup relationships between all lists" -ForegroundColor Gray
Write-Host "  • Recurring revenue models (Monthly, Quarterly, etc.)" -ForegroundColor Gray
Write-Host "  • Random recurrence counts (1-60 occurrences)" -ForegroundColor Gray
Write-Host "  • Start dates calculated 5-90 days after close" -ForegroundColor Gray
Write-Host "  • Stage-appropriate milestones with realistic timelines" -ForegroundColor Gray

# Disconnect
try {
    Disconnect-PnPOnline
    Write-Host ""
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
} catch {
    # Ignore disconnect errors
}
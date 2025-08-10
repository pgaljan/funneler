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
            
            $opportunity = @{
                OpportunityName = $opportunityTemplates[$i]
                Status = $status
                OpportunityStage = $stageAssignments[$i]
                Amount = $amount
                Probability = $probability
                OpportunityOwner = $owners | Get-Random
                Close = Get-RandomWeekday -DaysFromNow 512
                NextMilestone = $milestones | Get-Random
                NextMilestoneDate = (Get-Date).AddDays((Get-Random -Minimum 5 -Maximum 60))
            }
            
            # Assign customer based on our distribution plan
            $customerIndex = $customerAssignments[$i]
            if ($customerIndex -ge $customers.Count) {
                $customerIndex = $i % $customers.Count
            }
            $customer = $customers[$customerIndex]
            $opportunity[$customerFieldName] = $customer.Id
            
            # Check if opportunity already exists
            $existingOpp = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='OpportunityName'/><Value Type='Text'>$($opportunity.OpportunityName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
            
            if (-not $existingOpp) {
                Add-PnPListItem -List $ListName -Values $opportunity -ErrorAction Stop
                Write-Host "  ✓ Added opportunity: $($opportunity.OpportunityName) ($($opportunity.Amount.ToString('C0')))" -ForegroundColor Green
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
    Write-Host "✓ Customers list: $totalCustomersAdded new records added (30 total)" -ForegroundColor Gray
    Write-Host "  Distribution: Various industries and company types" -ForegroundColor Gray
}

if ($opportunitiesExists) {
    Write-Host "✓ Opportunities list: $totalOpportunitiesAdded new records added (40 total)" -ForegroundColor Gray
    Write-Host "  Value range: $25,000 - $2,500,000 (normal distribution)" -ForegroundColor Gray
    Write-Host "  Probability: 30% Low, 50% Medium, 20% High" -ForegroundColor Gray
    Write-Host "  Status: 80% Active, 15% At Risk, 5% Critical" -ForegroundColor Gray
    Write-Host "  Customer distribution: 1 customer (5 opps), 2 customers (3 opps each)," -ForegroundColor Gray
    Write-Host "                         5 customers (0 opps), remaining distributed" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Sample data includes:" -ForegroundColor Cyan
Write-Host "  • 30 diverse customers across multiple industries" -ForegroundColor Gray
Write-Host "  • 40 opportunities with realistic value distribution" -ForegroundColor Gray
Write-Host "  • Normal distribution of opportunity values ($25K-$2.5M)" -ForegroundColor Gray
Write-Host "  • Varied opportunity stages and probability levels" -ForegroundColor Gray
Write-Host "  • Strategic customer-opportunity assignments" -ForegroundColor Gray
Write-Host "  • Realistic contact information and milestone dates" -ForegroundColor Gray
Write-Host "  • Proper lookup relationships between lists" -ForegroundColor Gray

# Disconnect
try {
    Disconnect-PnPOnline
    Write-Host ""
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
} catch {
    # Ignore disconnect errors
}
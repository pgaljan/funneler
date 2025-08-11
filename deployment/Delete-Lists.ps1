param(
    [string]$SiteUrl,
    [string]$ClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e",
    [switch]$Force,
    [string[]]$ListNames = @(),
    [string]$DeletePrefix = ""
)

# If parameters not provided via command line, prompt for them
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    Write-Host "=== Quick SharePoint List Manager ===" -ForegroundColor Cyan
    Write-Host ""
    
    $SiteUrl = Read-Host "Enter SharePoint Site URL"
    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        Write-Error "Site URL is required. Exiting."
        exit 1
    }
}

# Allow override of default ClientId if not specified via command line
if (-not $PSBoundParameters.ContainsKey('ClientId')) {
    $defaultClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e"
    $ClientIdInput = Read-Host "Enter Client ID (press Enter for default: $defaultClientId)"
    if (-not [string]::IsNullOrWhiteSpace($ClientIdInput)) {
        $ClientId = $ClientIdInput
    }
}

Write-Host ""
Write-Host "=== Quick SharePoint List Manager ===" -ForegroundColor Cyan
Write-Host "Site: $SiteUrl" -ForegroundColor Yellow
Write-Host ""

# Import and connect
try {
    Import-Module PnP.PowerShell -Force -ErrorAction Stop
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
    $web = Get-PnPWeb
    Write-Host "✓ Connected to: $($web.Title)" -ForegroundColor Green
} catch {
    Write-Error "Connection failed: $($_.Exception.Message)"
    exit 1
}

# Get lists quickly - only custom lists, no libraries or system lists
Write-Host "Retrieving lists..." -ForegroundColor Yellow
try {
    $lists = Get-PnPList | Where-Object { 
        $_.BaseTemplate -eq 100 -and  # Generic/Custom Lists only
        -not $_.Hidden 
    } | Sort-Object Title
    
    Write-Host "✓ Found $($lists.Count) custom lists" -ForegroundColor Green
} catch {
    Write-Error "Failed to retrieve lists: $($_.Exception.Message)"
    exit 1
}

if ($lists.Count -eq 0) {
    Write-Warning "No custom lists found."
    Disconnect-PnPOnline
    exit 0
}

# Display lists in table format
Write-Host ""
Write-Host "Lists:" -ForegroundColor Cyan
$tableData = @()
for ($i = 0; $i -lt $lists.Count; $i++) {
    $list = $lists[$i]
    $tableData += [PSCustomObject]@{
        '#' = $i + 1
        'List Name' = $list.Title
        'Created' = $list.Created.ToString("yyyy-MM-dd")
        'Modified' = $list.LastItemModifiedDate.ToString("yyyy-MM-dd HH:mm")
        'Owner' = $list.Author.Title
    }
}

$tableData | Format-Table -AutoSize

# Handle command-line specified lists or prefix
if ($ListNames.Count -gt 0) {
    Write-Host "Processing specified lists: $($ListNames -join ', ')" -ForegroundColor Yellow
    $selectedLists = $lists | Where-Object { $_.Title -in $ListNames }
    
    if ($selectedLists.Count -eq 0) {
        Write-Warning "None of the specified lists were found."
        Disconnect-PnPOnline
        exit 1
    }
} elseif (-not [string]::IsNullOrWhiteSpace($DeletePrefix)) {
    Write-Host "Processing lists with prefix: $DeletePrefix" -ForegroundColor Yellow
    $selectedLists = $lists | Where-Object { $_.Title -like "$DeletePrefix*" }
    
    if ($selectedLists.Count -eq 0) {
        Write-Warning "No lists found with prefix '$DeletePrefix'."
        Disconnect-PnPOnline
        exit 1
    }
    
    Write-Host "Found $($selectedLists.Count) lists matching prefix:" -ForegroundColor Green
    $selectedLists | ForEach-Object { Write-Host "  • $($_.Title)" -ForegroundColor Yellow }
} else {
    # Interactive selection
    Write-Host ""
    Write-Host "Selection Options:" -ForegroundColor Yellow
    Write-Host "  Numbers: 1,3,5 or 1-5,8 or all" -ForegroundColor Gray
    Write-Host "  Prefix:  prefix:crm (deletes all lists starting with 'crm')" -ForegroundColor Gray
    Write-Host "  Cancel:  quit/exit/none" -ForegroundColor Gray
    Write-Host ""
    
    $selection = Read-Host "Enter selection"
    
    if ([string]::IsNullOrWhiteSpace($selection) -or $selection.ToLower() -in @("quit", "exit", "none", "cancel")) {
        Write-Host "Cancelled." -ForegroundColor Gray
        Disconnect-PnPOnline
        exit 0
    }
    
    # Handle prefix selection
    if ($selection.ToLower().StartsWith("prefix:")) {
        $prefix = $selection.Substring(7).Trim()
        $selectedLists = $lists | Where-Object { $_.Title -like "$prefix*" }
        
        if ($selectedLists.Count -eq 0) {
            Write-Warning "No lists found with prefix '$prefix'."
            Disconnect-PnPOnline
            exit 0
        }
        
        Write-Host "Found $($selectedLists.Count) lists with prefix '$prefix':" -ForegroundColor Green
        $selectedLists | ForEach-Object { Write-Host "  • $($_.Title)" -ForegroundColor Yellow }
    } elseif ($selection.ToLower() -eq "all") {
        $selectedLists = $lists
    } else {
        # Parse numeric selection
        try {
            $selectedIndices = @()
            $parts = $selection -split ','
            
            foreach ($part in $parts) {
                $part = $part.Trim()
                if ($part -match '^\d+$') {
                    $index = [int]$part - 1
                    if ($index -ge 0 -and $index -lt $lists.Count) {
                        $selectedIndices += $index
                    }
                } elseif ($part -match '^\d+-\d+$') {
                    $range = $part -split '-'
                    $start = [int]$range[0] - 1
                    $end = [int]$range[1] - 1
                    if ($start -ge 0 -and $end -lt $lists.Count -and $start -le $end) {
                        $selectedIndices += $start..$end
                    }
                }
            }
            
            $selectedIndices = $selectedIndices | Sort-Object | Get-Unique
            $selectedLists = @()
            foreach ($index in $selectedIndices) {
                $selectedLists += $lists[$index]
            }
            
            if ($selectedLists.Count -eq 0) {
                Write-Warning "No valid lists selected."
                Disconnect-PnPOnline
                exit 0
            }
        } catch {
            Write-Error "Invalid selection format."
            Disconnect-PnPOnline
            exit 1
        }
    }
}

# Confirmation
Write-Host ""
Write-Host "Selected for deletion:" -ForegroundColor Red
$selectedLists | ForEach-Object { Write-Host "  • $($_.Title)" -ForegroundColor Yellow }

Write-Host ""
Write-Host "WARNING: This cannot be undone!" -ForegroundColor Red -BackgroundColor Yellow

if (-not $Force) {
    $confirm = Read-Host "Type 'DELETE' to confirm"
    if ($confirm -ne "DELETE") {
        Write-Host "Cancelled." -ForegroundColor Gray
        Disconnect-PnPOnline
        exit 0
    }
}

# Delete lists
Write-Host ""
Write-Host "Deleting lists..." -ForegroundColor Yellow
$success = 0
$failed = 0

foreach ($list in $selectedLists) {
    try {
        Write-Host "Deleting '$($list.Title)'..." -NoNewline
        Remove-PnPList -Identity $list.Id -Force -ErrorAction Stop
        Write-Host " ✓" -ForegroundColor Green
        $success++
    } catch {
        Write-Host " ✗ $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

Write-Host ""
Write-Host "Results: $success deleted, $failed failed" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Yellow" })

# Cleanup
Disconnect-PnPOnline
Write-Host "Done." -ForegroundColor Cyan
This workflow brings you through the process of an automatic deployment of the sales funneler solution.


## Prerequisites

Install Powershell 7
```powershell
winget install --id Microsoft.Powershell --source winget
```

Run as Administrator
```powershell
Start-Process pwsh -Verb RunAs
```

Configure Execution Policy
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
Get-ExecutionPolicy -List
```

Install PnP PowerShell module
```powershell
Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
Get-Module -Name PnP.PowerShell -ListAvailable
mport-Module PnP.PowerShell -Force
```

Configure TLS/Security
```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
```

Verify Access
```powershell
$siteUrl = "https://your-tenant-name.sharepoint.com/sites/your-site-name"
Connect-PnPOnline -Url $siteUrl -Interactive
```

Authentication methods will vary by site.  
![Auth Error](./images/autherror.png)

If you get an auth error like this, create an app registration in entra, with redirect URIs as below:
![Redirect](./images/redirect.png)

## List configuration

The following scripts leverage interactive authentication, meaning you will be prompted for the modern dev click through authentication experience at each step, and giving you the opportunity to observe the changes to the sharepoint site.  If you do not pass arguments, it will prompt you for a Sharepoint Site name and a deployment prefix that will be prepended to the name of each table. 

![deploy scripts](./images/deployScripts.png)
> The excel file autoforms your powershell according to the excel file's configuration.  Highly recommend you leverage this

### 1. Create List
```powershell
.\Deploy-Lists.ps1 -SiteUrl "https://your-tenant-name.sharepoint.com/sites/your-site-name/" -ListPrefix "yourPrefix" 
```
> Creates the core tables

### 2. Flag Fields as required
```powershell
.\Set-Required-Fields.ps1 -SiteUrl "https://your-tenant-name.sharepoint.com/sites/your-site-name/" -ListPrefix "yourPrefix" 
```
> Flags fields as required in the core tables

### 3. Add Comment Log
```powershell
.\Add-Comment-Log.ps1 -SiteUrl "https://your-tenant-name.sharepoint.com/sites/your-site-name/" -ListPrefix "yourPrefix" 
```
> Turns on List versioning for the Opportunities list, adds comment log

### 4. Populate list with sample data
```powershell
.\Add-Comment-Log.ps1 -SiteUrl "https://your-tenant-name.sharepoint.com/sites/your-site-name/" -ListPrefix "yourPrefix" 
```
> Optional, but useful for testing


#### 5. Clean up
- Adjust columns in views
- Apply [form body json](./form-body-json.md)
- Create selection pill colors
- Clear the lists

#### 6. Launch
- Refresh the Excel, verify function
- Deploy Power BI dashboard, verify function
- Add users to site
- Share PBI with regular users
- Copy excel for power users


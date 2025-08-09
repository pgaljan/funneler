

Enable Templating:
```powershell
Connect-SPOService -Url https://[tenant-name]-admin.sharepoint.com
Set-SPOSite https://[tenant-name].sharepoint.com/sites/SalesFunnel -DenyAddAndCustomizePages 0
```

To get this to work, I needed to create amn Azure App registration and leverage the ClientID.  You may not have to.
Success using powershell 7 as administrator for PnP
Test Connection
```powershell
# See if you can connect using the built in app registration
Connect-PnPOnline -Url "https://[tenant-name].sharepoint.com/sites/[site-name]/" -Interactive -ClientId "31359c7f-bd7e-475c-86db-fdb8c937548e"
```

Test Site Access
```powershell
Get-PnPWeb
```
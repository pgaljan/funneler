# About


## Quick Start (Excel)
1. Identify the Site URL and List Prefix of your pipeline

```
https://contoso.sharepoint.com/sites/Sales/Lists/crmCustomers/
```
> tenant: **contoso**; 
> site: **Sales**;
> prefix: **crm**


2. Excel Configuration

Open `Sales Funnel Sharepoint.xlsx` and navigate to **Settings**

![List Specification](./images/listSelect.png)
>Specify the SiteURL and prefix

1. Refresh and test links

![Refresh](./images/refresh.png)

1. (Optional) Configure refresh
Refresh settings are the defaults for Excel.  If using in production, consider adding a refresh on open, and clearing data on refresh.

>
## Screenshots

### Pipeline Dashboard
![Pipeline Dashboard](./images/dashboard.png)
*Main dashboard showing opportunities, stages, and key metrics*

### SharePoint Integration

*List View in Sharepoint*
![Sharepoint List View](./images/opportunityList.png)

*Calendar view in Sharepoint*
![SharePoint Calendar View](./images/calendarview.png)

*Opportunity form*
![SharePoint Opportunity Form](./images/opportunityform.png)

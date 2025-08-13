# Sharepoint Sales Funneler

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![SharePoint](https://img.shields.io/badge/SharePoint-Online-blue.svg)](https://www.microsoft.com/sharepoint)
[![Excel](https://img.shields.io/badge/Excel-365-green.svg)](https://www.microsoft.com/excel)
[![Power Query](https://img.shields.io/badge/Power%20Query-Enabled-orange.svg)](https://powerquery.microsoft.com/)
[![PowerShell 7 Required](https://img.shields.io/badge/PowerShell%207-Recommended-yellow.svg)](https://github.com/PowerShell/PowerShell)

> A Sharepoint and Excel learning kit comprised of a production-grade sales pipeline management system for small to mid-sized teams.  Featuring  on-demand sync with Excel for visualization, and containing thorough documentation, it is a solid introduction to building a sustainable, governable two-tier application in the Microsoft ecosystem. The simple, three-table data model allows the exploration of these concepts without requiring a background in data structure management, while maintaining usefulness and stability when deployed in production.


## Documentation Quick Start
- [DevOps Guide](./docs/auto-deploy.md)
- [Citizen Developer Guide](./docs/manual-deploy.md)
- [Security & Governance](./docs/security.md)
- [Code Guide](./docs/code-guide.md)
- [Form Body JSON](./docs/form-body-json.md)

## Operator Features
**SharePoint DevOps** will enjoy:

- **Easy Governance Controls** - Protect sensitive data and comply with regulations using existing M365 governance policies
- **Flexible Deployment Methods** - Deploy in minutes with near-complete [powershell automation](./docs/auto-deploy.md)
- **Thorough Documentation** - Read the [code guide](./docs/code-guide.md) to understand the ground-up implementation
- **Integrated IAM** - Use regular M365 features for sharing, user tagging and commenting
- **Bulk list management** - Manage lists en masse
- **Auditing tools** - Audit and drift detection scripts
- **Worry-free Licensing** - MIT License

## User Features
**Business Developers** will find a robust set of pipeline management features: 

- **Dynamic Pipeline Dashboard** - Stateless, visually rich sales funnel
- **Risk Reporting** - sales phase management and alerting
- **Fiscal Year Support** - Customizable fiscal quarters and calendar systems
- **Milestone Tracking** - Next steps and deadline management
- **Visual Status Indicators** - At-a-glance opportunity health
- **Hyperlinked Navigation** - Direct links to SharePoint records
- **Multi-user Collaboration** - SharePoint-backed team workflows

## Project builder experience
**Citizen developers** deploying on their own will learn: 
- **Excel functions** - Excel formulas, conditions, logic, arrays
- **Dashboarding** - Visualizations, conditional formatting, array presentation
- **Data Modeling** - Data Entity Relationships
- **Extract/Transform/Load** - Environment-portable ETL using Power Query
- **SharePoint Automation** - PowerShell 7 PnP automation & authentication
- **SharePoint Governance** - Audit, drift reporting, DLP and governance controls

# Quick Start (Excel)
1. Identify the Site URL and List Prefix of your pipeline

```
https://contoso.sharepoint.com/sites/Sales/Lists/crmCustomers/
```
> tenant: **contoso**; 
> site: **Sales**;
> prefix: **crm**


2. Excel Configuration

Open `Sales Funnel Sharepoint.xlsx` and navigate to **Settings**

![List Specification](./docs/images/listSelect.png)
>Specify the SiteURL and prefix

3. Refresh and test links

![Refresh](./docs/images/refresh.png)

4. (Optional) Configure refresh
Refresh settings are the defaults for Excel.  If using in production, consider adding a refresh on open, and clearing data on refresh.

## Architecture
```mermaid
graph TD
    B[Power BI]
    C[Excel]
    A[SharePoint<br/>• Customers<br/>• Opportunities<br/>• Milestones] 
    
    B -.->|REST API| A
    C -.->|REST API| A
```

## Relationship Diagram

```mermaid
erDiagram
    CUSTOMERS ||--o{ OPPORTUNITIES : "has"
    OPPORTUNITIES ||--o{ MILESTONES : "contains"
    OPPORTUNITIES ||--o{ TRANSACTIONS : "contains"
    
    CUSTOMERS {
        int CustomerId PK "Primary identifier"
        string CustomerName 
        string Website
        string NAICScode "[]"
        string Status "[]"
    }
    
    OPPORTUNITIES {
        int OpportunityId PK "Primary identifier"
        string OpportunityName 
        string Status "[]"
        string OpportunityOwner
        string Stage "[]"
        currency OpportunityValue
        string WinProbability "[]"
        datetime ExpectedCloseDate
        string RecurringRevenueModel "[]"
        number Recurrences
        datetime Start_Date
        int CustomerId FK 
        text Comment_Log "multi-line append" 
    }
    
    MILESTONES {
        int MilestoneId PK "Milestone name"
        string MilestoneName
        string OpportunityId FK 
        string Owner "Person responsible"
        datetime Date "Milestone date"
        string Status "[]"
    }

    TRANSACTIONS {
        int TransactionId PK 
        int OpportunityId FK
        int CustomerId FK
        datetime TransactionDate 
        string Status "[]"
    }

```
> `Transactions` table is calculated via power query or Lambda function, depending on implementation
>
## Screenshots

### Pipeline Dashboard
![Pipeline Dashboard](docs/images/dashboard.png)
*Main dashboard showing opportunities, stages, and key metrics*

### SharePoint Integration

*List View in Sharepoint*
![Sharepoint List View](docs/images/opportunityList.png)

*Calendar view in Sharepoint*
![SharePoint Calendar View](docs/images/calendarview.png)

*Opportunity form*
![SharePoint Opportunity Form](docs/images/opportunityform.png)




## Requirements

- **Microsoft 365** with SharePoint Online
- **Excel 365** with Power Query support
- **SharePoint Site** with list creation permissions


## Security & Permissions

### SharePoint Permissions
- CRUD permissions based on sharepoint list attributes

### Data Protection
- All data stored in Microsoft 365 tenant
- Inherits organizational security policies
- Audit trails available through SharePoint
- GDPR compliance through Microsoft 365

## Performance & Scalability

### Current Limits 
*(effectively tied to Sharepoint list scalability)*
- **Opportunities**: 5,000 items (recommended)
- **Concurrent Users**: ~50 users per list

### Scaling Recommendations
- Archive closed opportunities annually
- Consider dedicated SharePoint sites for scaleout and refined RBAC segmentation

## Known Issues

## Roadmap
- PBI dashboard
- Recurring revenue setup
- User-defined Phase
- User defined Status
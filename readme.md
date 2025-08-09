# Sharepoint Sales Funneler

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![SharePoint](https://img.shields.io/badge/SharePoint-Online-blue.svg)](https://www.microsoft.com/sharepoint)
[![Excel](https://img.shields.io/badge/Excel-365-green.svg)](https://www.microsoft.com/excel)
[![Power Query](https://img.shields.io/badge/Power%20Query-Enabled-orange.svg)](https://powerquery.microsoft.com/)
[![PowerShell 7 Required](https://img.shields.io/badge/PowerShell%207-Recommended-yellow.svg)](https://github.com/PowerShell/PowerShell)

> A Sharepoint learning kit comprised of a production-grade sales pipeline management system for small to mid-sized teams.  Featuring  on-demand sync with Excel for visualization, and containing thorough documentation on the visualization, ETL, and list deployment processes. Anyone building this project will exercise skills with excel formulas, data visualization and conditional formatting, array handling, power query.  Those interested in learning and practicing automated deployment methods can leverage the included powershell PnP deployment scripts.


## Features

- **Easy Governance Controls** - Protect sensitive data and comply with regulations using existing M365 governance policies
- **Flexible Deployment Methods** - Deploy manually using the [citizen developer](./docs/self-deploy.md) docs or the included PowerShell 7 scripts
- **Integrated IAM** - Use regular M365 features for sharing, user tagging and commenting
- **Dynamic Pipeline Dashboard** - Stateless, visually rich sales funnel in Excel and PBI (coming soon)
- **SharePoint Integration** - On-demand uni-directional sync with user-defined SharePoint lists
- **Fiscal Year Support** - Customizable fiscal quarters and calendar systems
- **Milestone Tracking** - Next steps and deadline management
- **Visual Status Indicators** - At-a-glance opportunity health
- **Hyperlinked Navigation** - Direct links to SharePoint records
- **Multi-user Collaboration** - SharePoint-backed team workflows

## Screenshots

### Pipeline Dashboard
![Pipeline Dashboard](docs/images/dashboard.png)
*Main dashboard showing opportunities, stages, and key metrics*

### SharePoint Integration
![Sharepoint List View](docs/images/opportunityList.png)
*List View in Sharepoint*
![SharePoint Calendar View](docs/images/calendarview.png)
*Calendar view in Sharepoint*
![SharePoint Opportunity Form](docs/images/opportunityform.png)
*Opportunity form*


## Architecture

```mermaid
graph TD
Opportunities --> PowerQuery[Power Query] 
Customers --> PowerQuery
PowerQuery --> Excel[Excel/PBI]
Excel -.->|hyperlink| Opportunities
Excel -.->|hyperlink| Customers
```

## Requirements

- **Microsoft 365** with SharePoint Online
- **Excel 365** with Power Query support
- **SharePoint Site** with list creation permissions

## Documentation
- [admin guide](./docs/admin-guide.md): Automated deployment and security evaluation of the funneler solution
- [self-deploy](./docs/self-deploy.md): Deployment guide for the citizen developer
- [code-guide](./docs/code-guide.md): information about the ETL process and a breakdown of excel functions leveraged
- [security](./docs/security.md): information about security

## Repository Structure

```
funneler/
├── Sales Funnel.xlsx       # Main Excel dashboard
├── templates/
│   ├── opportunities.stp   # Sharepoint template
│   ├── customers.stp       # Sharepoint template
│   └── template.xlsx       # Excel template
├── deployment/
│   ├── SalesFunnel.xml     # PnP provisioning template
│   ├── deploy.ps1          # Deployment script
│   └── permissions.ps1     # Security configuration
├── power-query/
│   ├── opportunities.m     # Opportunities data source
│   └── customers.m         # Customers data source
├── docs/
│   ├── images/
│   ├── admin-guide.md      # Automated deployment
│   ├── security.md         # Security assessment script, considerations & reading
│   ├── code-guide.md       # m code and excel formulas
│   └── self-deploy.md      # Citizen developer guide for deployment
└── README.md
```

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
- comment log disappears from excel power query, but remains in list (versioning limitation)

## Roadmap
- M code refactor to resolve commment log issue
- PBI dashboard
- Expand customer metadata
- Recurring revenue setup
- User-defined Phase
- User defined Status
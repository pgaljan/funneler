# Sharepoint Sales Funneler

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![SharePoint](https://img.shields.io/badge/SharePoint-Online-blue.svg)](https://www.microsoft.com/sharepoint)
[![Excel](https://img.shields.io/badge/Excel-365-green.svg)](https://www.microsoft.com/excel)
[![Power Query](https://img.shields.io/badge/Power%20Query-Enabled-orange.svg)](https://powerquery.microsoft.com/)
[![PowerShell 7 Required](https://img.shields.io/badge/PowerShell%207-Recommended-yellow.svg)](https://github.com/PowerShell/PowerShell)

> A Sharepoint and Excel learning kit comprised of a production-grade sales pipeline management system for small to mid-sized (50 user) teams.  Featuring  on-demand sync with Excel for visualization, and containing thorough documentation, it is a solid introduction to building a sustainable, governable two-tier application in the Microsoft ecosystem. The simple, three-table data model allows the exploration of these concepts without requiring a background in data structure management, while maintaining usefulness and stability when deployed in production.



## Documentation Quick Start
- [DevOps Guide](./docs/auto-deploy.md)
- [Citizen Developer Guide](./docs/manual-deploy.md)
- [BizOps Guide](./docs/bizops-guide.md)
- [Security & Governance](./docs/security.md)
- [Code Guide](./docs/code-guide.md)
- [Form Body JSON](./docs/form-body-json.md)

## Features


**Business Operations Professionals** will find a robust set of pipeline management features: 

- **Dynamic Pipeline Dashboard** - Stateless, visually rich sales funnel
- **Risk Management** - Quarterly Revenue-to-Risk calculations
- **Fiscal Year Support** - Customizable fiscal quarters and calendar systems
- **Milestone Tracking** - Deliverable management
- **Document Library** - Attach any document type to any record in the pipeline
- **Commenting and user tagging** - Office-style comments, user tagging, and actions 
- **Opportunity Health** - User-defined, score-based opportunity health indicators 
- **Hyperlinked Navigation** - Direct links to SharePoint records
- **Multi-user Collaboration** - SharePoint-backed team workflows
- **Flexible Consumption Model** - Stateless Excel or Power BI frontend

**SharePoint Developer/Operators** will enjoy:

- **Easy Governance Controls** - Protect sensitive data and comply with regulations using existing M365 governance policies
- **Flexible Deployment Methods** - Deploy from code in minutes with near-complete [powershell automation](./docs/auto-deploy.md)
- **Thorough Documentation** - Read [DevOps](./docs/auto-deploy.md) and [code guide](./docs/code-guide.md) to understand the ground-up implementation
- **Integrated IAM** - Use regular M365 features for self-service or workflow-driven user management
- **Bulk list management** - [Manage lists](./deployment/day-2/readme.md/#2-manage-listsps1) en masse
- **Auditing tools** - [Audit](./deployment/day-2/readme.md/#1-audit-list-securtyps1) and drift detection scripts
- **Worry-free Licensing & Roadmap** - [MIT License](./LICENSE), openly contributable [issue backlog](https://github.com/pgaljan/funneler/issues), and robust [Code of Conduct](CODE_OF_CONDUCT.md) framework

**Project Builders** deploying on their own will learn: 
- **Excel functions** - [Excel formulas](./docs/code-guide.md/#formula-patterns), conditions, logic, arrays
- **Ready-made dataset** - realistic sample data to test out modeling, visualization and summarization techniques
- **Dashboarding** - [Visualizations](#screenshots), conditional formatting, array presentation
- **Data Modeling** - [Data entity](#relationship-diagram) relationships
- **Extract/Transform/Load** - [Environment-portable ETL](./docs/code-guide.md/#power-query-m-code) using Power Query
- **Form Construction** - Form customization with [JSON](./docs/form-body-json.md)
- **SharePoint Automation** - PowerShell 7 PnP [automation](./docs/auto-deploy.md#1-create-list) & [authentication](./docs/auto-deploy.md/#prerequisites)
- **SharePoint Governance** - [Security](./docs/security.md), audit, drift reporting, DLP and governance controls

## Getting Started 
- BizOps guide is under construction, refer to the [Quick Start](./docs/bizops-guide.md/#quick-start-excel)

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


## Issues

![Issues](https://img.shields.io/github/issues/pgaljan/funneler)
![Closed Issues](https://img.shields.io/github/issues-closed/pgaljan/funneler)
![Pull Requests](https://img.shields.io/github/issues-pr/pgaljan/funneler)

### Filing Issues
Leverage the [user persona](./docs/user-persona.md) to create user stories around bug and feature issues.  Clear use cases will be prioritized, so be sure to fill out as many template prompts as you are able.
  

### Known Issues
- **[Issue 7](https://github.com/pgaljan/funneler/issues/7)** -  ETL will not work across DevOps and User-directed deployment 
  - (*Workaround* - mutliple dashboards)


### Roadmap
| DevOps | BizOps | Learner |
| --- | --- | --- |
| [One-command install](https://github.com/pgaljan/funneler/issues/1) |[BizOps User Guide](https://github.com/pgaljan/funneler/issues/6) | [Text Walkthroughs](https://github.com/pgaljan/funneler/issues/15) |
| [Deploy Versioning](https://github.com/pgaljan/funneler/issues/18) | [Converged ETL](https://github.com/pgaljan/funneler/issues/7) | [Video Walkthroughs](https://github.com/pgaljan/funneler/issues/16) |
| | [PBI dashboard](https://github.com/pgaljan/funneler/issues/4) |  |
| | [Mobile enablement](https://github.com/pgaljan/funneler/issues/17) |  |
| | [Cascading Milestone selector](https://github.com/pgaljan/funneler/issues/10) |  |
|  | [User-defined Attributes](https://github.com/pgaljan/funneler/issues/11) |  |


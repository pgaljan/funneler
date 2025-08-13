### Manual Deployment
Use this guide if you would like to deploy this pipeline as a citizen developer, do not have access to or would otherwise prefer to avoid the powershell required to perform automated list deployment.

If you plan to deploy more than a handful of these sites, I recommend you investigate the automated methods documented in the [admin guide](./admin-guide.md).  Setting this up without the use of templates is tedious, and realistically takes about 15 minutes.  An automated deploy is available in moments, without nearly as many tweaks needed, and can be integrated into automated, helpdesk-triggered workflows.

In either case, it is **strongly** recommended that maintain your RBAC permissions at the site level (as opposed to list level). To maintain close and segmented control of permissions, consider deploying each sales pipeline to its own dedicated sharepoint site.  If you choose to deploy multiple sales pipelines to the same sharepoint site, please consider leveraging [automated deployment](./admin-guide.md) methods to ensure the lists are secured according to your business requirements.

Before deploying on your own, read the short [security guide](./security.md) to learn or refresh your memory around regulatory compliance requirements for systems like this.

Refer to the [code guide](./code-guide.md) if you are interested in adapting the m code to other use cases or learning more about the formulas leveraged in visualizations.

- [Manual Deployment](#manual-deployment)
- [Process](#process)
  - [1. Determine your Tenant Name](#1-determine-your-tenant-name)
  - [2. Locate or create a Sharepoint Team Site](#2-locate-or-create-a-sharepoint-team-site)
  - [3. Create lists](#3-create-lists)
  - [4. Create Lookups](#4-create-lookups)
  - [5. Clean up](#5-clean-up)
  - [6. Launch](#6-launch)

### Process

#### 1. Determine your Tenant Name

Before you proceed, you will need your sharepoint **tenant name** and a sharepoint **site name**

To determine your **tenant name**, navigate to <a href="https://portal.office.com" target="_blank">portal.office.com</a> and select "Apps" in the left navigation.  Click on "Sharepoint".  The string before sharepoint.com is your tenant name.  In the following example it is "yourPortalName".

https://**yourPortalName**.sharepoint.com/_layouts/15/sharepoint.aspx?&login_hint=yourUserNameg@yourOrgDomain.com

#### 2. Locate or create a Sharepoint Team Site 

Refer to [this video](https://www.youtube.com/embed/HQw5nRwAJFc?si=lQHoK6gRMOGDAvXW) for more information about creating SharePoint sites.

#### 3. Create lists 
Use a consistent available prefix with the following templates
> **IMPORTANT**:  The excel and PBI templates assume a consistent naming convention of **prefixListname**.  If you don't follow the convention, the record linking will not work

- [Customers](../deployment/templates/Customers.csv) 
- [Opportunities](../deployment/templates/Opportunities.csv)
- [Milestones](../deployment/templates/Milestones.csv)

#### 4. Create Lookups
- Opportunities > Customers
- Milestones > Opportunities

#### 5. Clean up
- Adjust columns in views
- Apply [form body json](./form-body-json.md)
- Create selection pill colors
- Create test opportunities

#### 6. Launch
- Refresh the Excel, verify function
- Deploy Power BI dashboard, verify function
- Add users to site
- Share PBI with regular users
- Copy excel for power users


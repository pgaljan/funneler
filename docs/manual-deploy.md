### Manual Deployment
Use this guide if you would like to deploy this pipeline as a citizen developer, do not have access to or would otherwise prefer to avoid the powershell required to perform automated list deployment.

If you plan to deploy more than a handful of these sites, I recommend you investigate the automated methods documented in the [admin guide](./admin-guide.md).  Setting this up without the use of templates is tedious, and realistically takes about 15 minutes.  An automated deploy is available in moments, without nearly as many tweaks needed, and can be integrated into automated, helpdesk-triggered workflows.

In either case, it is **strongly** recommended that maintain your RBAC permissions at the site level (as opposed to list level). To maintain close and segmented control of permissions, consider deploying each sales pipeline to its own dedicated sharepoint site.  If you choose to deploy multiple sales pipelines to the same sharepoint site, please consider leveraging [automated deployment](./admin-guide.md) methods to ensure the lists are secured according to your business requirements.

Before deploying on your own, read the short [security guide](./security.md) to learn or refresh your memory around regulatory compliance requirements for systems like this.

Refer to the [code guide](./code-guide.md) if you are interested in adapting the m code to other use cases or learning more about the formulas leveraged in visualizations.

### Process

#### Locate your Tenant Name

Before you proceed, you will need your sharepoint **tenant name** and a sharepoint **site name**

To determine your **tenant name**, navigate to <a href="https://portal.office.com" target="_blank">portal.office.com</a> and select "Apps" in the left navigation.  Click on "Sharepoint".  The string before sharepoint.com is your tenant name.  In the following example it is "yourPortalName".

https://**yourPortalName**.sharepoint.com/_layouts/15/sharepoint.aspx?&login_hint=yourUserNameg@yourOrgDomain.com

#### Locate or create a Sharepoint Team Site 

Refer to [this video](https://www.youtube.com/embed/HQw5nRwAJFc?si=lQHoK6gRMOGDAvXW) for more information about creating SharePoint sites.

#### Create lists 
Use a consistent available prefix
- [Customers](../deployment/Customers.csv)
- [Opportunities](../deployment/Opportunities.csv)

#### Create Custoner Name Lookup
> Copy/Paste from CustomerId


#### (optional) apply body JSON format
Leverage this as body JSON on the respective forms to clean up the look and feel for data entry.

If you plan to make extensive use of the sharepoint forms, you should apply [form body json](./form-body-json.md)

#### Prep the lists for prod
- Delete all rows
- Add a test customer
- Add a test opportunity
- Add users to the site members

#### Launch and modify settings in Citizen Deployed Frontend excel
- Open the excel
- Modify site URL and list names
- Refresh all
- Verify functionality

#### (Optional) Create and Link Calendar View
Lists has a very functional calendar view.  Navigate to the Opportunities list, and create it.  Note the view GUID and populate it in the settings.
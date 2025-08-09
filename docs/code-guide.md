## Power Query (M Code)
These queries provide a reusable template for connecting to SharePoint sites and lists defined by variables within the workbook itself. To implement this code, developers should create two named ranges in their Excel workbook: "siteUrl" containing the SharePoint site URL, and "custList" containing the exact name of the customer list as it appears in SharePoint. The query automatically handles SharePoint authentication, filters to the specified list, and expands the customer data into a flat table structure suitable for Power BI integration or Excel analysis and visualization. 

### Opportunities Data Connection
```m
let
    siteUrl = Excel.CurrentWorkbook(){[Name="siteUrl"]}[Content]{0}[Column1],
    listName = Excel.CurrentWorkbook(){[Name="oppList"]}[Content]{0}[Column1],
    Source = SharePoint.Tables(siteUrl, [Implementation=null, ApiVersion=15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Title] = listName)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Id", "Title"}),
    #"Expanded Items" = Table.ExpandTableColumn(#"Removed Columns", "Items", {"Id", "Title", "Status", "Stage", "Amount", "Probability", "CustomerId", "Close", "NextMilestone", "NextMilestoneDate", "Comment Log", "Customer"}, {"Id", "Title", "Status", "Stage", "Amount", "Probability", "CustomerId", "Close", "NextMilestone", "NextMilestoneDate", "Comment Log", "Customer"})
in
    #"Expanded Items"
```
### Customers Data Connection
```m
let
    siteUrl = Excel.CurrentWorkbook(){[Name="siteUrl"]}[Content]{0}[Column1],
    listName = Excel.CurrentWorkbook(){[Name="custList"]}[Content]{0}[Column1],
    Source = SharePoint.Tables(siteUrl, [Implementation=null, ApiVersion=15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Title] = listName)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Id", "Title"}),
    #"Expanded Items" = Table.ExpandTableColumn(#"Removed Columns", "Items", {"Id", "Title", "PrimaryContact", "PrimaryContactTitle", "AlternateContact", "AlternateContactTitle", "AlternateContact2", "AlternateContact2Title", "Website", "Modified", "Created"}, {"Id", "Title", "PrimaryContact", "PrimaryContactTitle", "AlternateContact", "AlternateContactTitle", "AlternateContact2", "AlternateContact2Title", "Website", "Modified", "Created"}),
    #"Expanded Website" = Table.ExpandRecordColumn(#"Expanded Items", "Website", {"Url"}, {"Url"})
in
    #"Expanded Website"
```

## Excel
### Formula Patterns

#### High Complexity Formulas
- **Dynamic Arrays**: Extensive use of UNIQUE, FILTER, SORTBY, and TRANSPOSE
- **Nested Lookups**: Multiple XLOOKUP functions with complex criteria
- **Structured References**: Heavy use of Excel table references (e.g., `View[[#This Row],[Id]]`)
- **Complex Text Operations**: Date formatting and URL generation

#### Primary Formula Patterns
1. **Data Lookup and Display**: XLOOKUP patterns for retrieving related data
2. **Dynamic Filtering**: FILTER and UNIQUE combinations for data analysis
3. **URL Generation**: HYPERLINK formulas for SharePoint integration
4. **Date Calculations**: Complex fiscal year and quarter calculations
5. **Data Visualization Support**: Array formulas for chart data preparation

#### External Dependencies
- SharePoint Lists (Opportunities, Customers)
- Named ranges for configuration settings
- External data connections for real-time updates


### Function Inventory

The spreadsheet uses 21 unique Excel functions:

| Function | Usage | Description |
|----------|-------|-------------|
| XLOOKUP | High | Advanced lookup function |
| HYPERLINK | High | Creates clickable links |
| SUBTOTAL | Medium | Subtotal calculations for dynamic filtering |
| UNIQUE | Medium | Returns unique values |
| TRANSPOSE | Medium | Transposes arrays |
| FILTER | Medium | Filters data arrays |
| SORTBY | Medium | Sorts data by criteria |
| IF | High | Conditional logic |
| MONTH | Medium | Extracts month from date |
| YEAR | Medium | Extracts year from date |
| RIGHT | Medium | Text extraction |
| CEILING | Low | Rounds up to nearest integer |
| ROUNDUP | Low | Rounds numbers up |
| LET | Low | Defines variables in formulas |
| SWITCH | Low | Multiple condition evaluation |
| TEXT | Low | Text formatting |
| CONCATENATE | Low | Text concatenation |
| SUM | Low | Summation |
| COUNT | Low | Count functions |
| MAX | Low | Maximum value |
| MIN | Low | Minimum value |
 

### Named Ranges Inventory

| Name | Type | Reference | Sheet | Hidden | Description |
|------|------|-----------|-------|---------|-------------|
| calendarType | User Defined | Pipeline!$B$2 | Global | No | Calendar type setting |
| custList | User Defined | Settings!$E$2 | Global | No | Customer list reference |
| dateRange | User Defined | Pipeline!$D$2 | Global | No | Date range setting |
| fqStart | User Defined | Settings!$B$1 | Global | No | Fiscal quarter start |
| oppList | User Defined | Settings!$E$3 | Global | No | Opportunities list reference |
| pipelineType | User Defined | Pipeline!$B$3 | Global | No | Pipeline type setting |
| probabilityThreshold | User Defined | Settings!$B$2 | Global | No | Probability threshold setting |
| siteUrl | User Defined | Settings!$E$1 | Global | No | SharePoint site URL |
| _xlchart.v2.0 | Chart Range | graphics!$A$11:$A$16 | Global | Yes | Chart data series 1 |
| _xlchart.v2.1 | Chart Range | graphics!$C$11:$C$16 | Global | Yes | Chart data series 2 |
| ExternalData_1 | External Data | 'Customers'!$A$1:$K$31 | Customers | Yes | Customer Sharepoint List |
| ExternalData_2 | External Data | Opportunities!$A$1:$L$37 | Opportunities | Yes | Opportunity Sharepoint List |
| ExternalData_3 | External Data | Pipeline!$M$7:$M$43 | Pipeline | Yes | Pipeline View |
| Slicer_Close_Quarter | Slicer | #N/A | Global | No | Close quarter slicer |
| Slicer_Stage | Slicer | #N/A | Global | No | Stage slicer |


